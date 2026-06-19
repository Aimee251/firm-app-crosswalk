"""
firm-app-crosswalk — full pipeline in ONE self-contained file.

    python run_pipeline.py            # run all three stages
    python run_pipeline.py --refilter # also rebuild the pilot file from the CSV

This file does NOT import Filter.py / AppClassifer.py / Webscapper_apple.py —
everything it needs is below, so stale copies of those won't break it.

Setup (once):
    pip install anthropic pandas requests openpyxl python-dotenv
    # AI_apis.env must contain:  ANTHROPIC_API_KEY=sk-ant-...

Every stage is resume-safe: rerun after any interruption and it picks up
where it left off.
"""

import os
import re
import sys
import json
import time
import pandas as pd
import requests
from urllib.parse import urlparse
from difflib import SequenceMatcher
from concurrent.futures import ThreadPoolExecutor, as_completed

import anthropic
from dotenv import load_dotenv

# ---------- file paths (the chain between stages) ----------
BASE = "/Users/tianyuzhou/Documents/Finance_RA"
CSV_FILE = f"{BASE}/pitchbook_wrds.csv"
PILOT_FILE = f"{BASE}/pitchbook_app_pilot.xlsx"
CLASSIFIED_FILE = f"{BASE}/pitchbook_app_classified.xlsx"
APPLE_FILE = f"{BASE}/apple_developer_apps.xlsx"

PILOT_SIZE = 1000
VALID_LABELS = ("YES", "NO", "UNCLEAR")


# =====================================================================
# STAGE 1 — Filter: first N companies + empty working columns
# =====================================================================
def stage_filter():
    df = pd.read_csv(
        CSV_FILE,
        usecols=["CompanyID", "CompanyName", "Description", "Website",
                 "PrimaryIndustryGroup", "PrimaryIndustryCode"],
    )
    pilot = df.head(PILOT_SIZE).copy()
    for col in ["Current_App_Status", "iOS_App_URL", "Android_app_URL",
                "Developer_page_URL", "Developer_Name", "Evidence_note",
                "Needs_Historical_review"]:
        pilot[col] = ""
    pilot.to_excel(PILOT_FILE, index=False)
    print(f"[filter] saved {len(pilot)} rows to {PILOT_FILE}")


# =====================================================================
# STAGE 2 — Classify with Claude Haiku (batched + parallel)
# =====================================================================
CLASSIFY_BATCH_SIZE = 20
CLASSIFY_WORKERS = 8
CLASSIFY_MODEL = "claude-haiku-4-5-20251001"
CLASSIFY_RETRIES = 5
CLASSIFY_SYSTEM = (
    "You classify whether a company CURRENTLY has its own mobile app(s), "
    "using its description, website, and industry. "
    "Reply with ONLY a JSON array — no markdown, no prose. "
    'One object per company: [{"index": <number>, "label": "YES"|"NO"|"UNCLEAR"}].'
)
_client = None  # set in stage_classify()


def _classify_message(batch):
    lines = []
    for idx, row in batch:
        desc = str(row.get("Description") or "")[:300]
        lines.append(
            f"[{idx}] Company: {row.get('CompanyName')} | "
            f"Website: {row.get('Website')} | "
            f"Industry: {row.get('PrimaryIndustryGroup')} "
            f"({row.get('PrimaryIndustryCode')}) | Description: {desc}"
        )
    return "\n".join(lines)


def _parse_labels(text):
    cleaned = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    out = {}
    for item in json.loads(cleaned):
        label = str(item.get("label", "")).strip().upper()
        if label in VALID_LABELS:
            out[int(item["index"])] = label
    return out


def _classify_batch(batch):
    msg = _classify_message(batch)
    for attempt in range(CLASSIFY_RETRIES):
        try:
            resp = _client.messages.create(
                model=CLASSIFY_MODEL,
                max_tokens=64 * len(batch) + 256,
                system=CLASSIFY_SYSTEM,
                messages=[{"role": "user", "content": msg}],
            )
            text = "".join(b.text for b in resp.content if b.type == "text")
            return _parse_labels(text)
        except (anthropic.RateLimitError, anthropic.APIStatusError) as e:
            if isinstance(e, anthropic.APIStatusError) and e.status_code not in (429, 500, 529):
                raise
            time.sleep(2 ** attempt * 5)
        except (json.JSONDecodeError, ValueError):
            return {}
    return {}


def stage_classify():
    global _client
    load_dotenv("AI_apis.env")
    _client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    if os.path.exists(CLASSIFIED_FILE):
        df = pd.read_excel(CLASSIFIED_FILE)
        print(f"[classify] resuming from {CLASSIFIED_FILE}")
    else:
        df = pd.read_excel(PILOT_FILE)
    if "AI_Result" not in df.columns:
        df["AI_Result"] = ""

    todo = [(i, row) for i, row in df.iterrows()
            if str(df.at[i, "AI_Result"]).strip().upper() not in VALID_LABELS]
    batches = [todo[s:s + CLASSIFY_BATCH_SIZE]
               for s in range(0, len(todo), CLASSIFY_BATCH_SIZE)]
    print(f"[classify] {len(todo)} companies, {len(batches)} batches, "
          f"{CLASSIFY_WORKERS} parallel")
    if not batches:
        print("[classify] nothing to do — all rows already labeled")
        df.to_excel(CLASSIFIED_FILE, index=False)
        return

    start, done = time.time(), 0
    with ThreadPoolExecutor(max_workers=CLASSIFY_WORKERS) as pool:
        futures = {pool.submit(_classify_batch, b): b for b in batches}
        for fut in as_completed(futures):
            result = fut.result()
            for idx, _ in futures[fut]:
                df.at[idx, "AI_Result"] = result.get(idx, "")
            done += 1
            if done % 5 == 0:
                df.to_excel(CLASSIFIED_FILE, index=False)
                print(f"[classify]   {done}/{len(batches)} batches done")

    df.to_excel(CLASSIFIED_FILE, index=False)
    print(f"[classify] done in {time.time() - start:.0f}s -> {CLASSIFIED_FILE}")
    print(df["AI_Result"].value_counts())


# =====================================================================
# STAGE 3 — Apple lookup via iTunes Search API (parallel, gentle)
# =====================================================================
APPLE_PROCESS = {"YES", "UNCLEAR"}
APPLE_WORKERS = 4          # keep small — Apple throttles the iTunes API
APPLE_CHECKPOINT = 25
APPLE_SEARCH_LIMIT = 10
APPLE_RETRIES = 4
NAME_NOISE = {
    "inc", "incorporated", "llc", "ltd", "limited", "corp", "corporation",
    "co", "company", "group", "holdings", "plc", "gmbh", "sa", "ag",
    "technologies", "technology", "labs", "lab", "software", "the", "app",
}


def _core_name(name):
    if not isinstance(name, str):
        return ""
    n = re.sub(r"[^a-z0-9 ]", " ", name.lower())
    return " ".join(t for t in n.split() if t and t not in NAME_NOISE).strip()


def _registered_domain(url):
    if not isinstance(url, str) or not url.strip():
        return ""
    u = url.strip()
    if "://" not in u:
        u = "http://" + u
    netloc = urlparse(u).netloc.lower().split(":")[0]
    if netloc.startswith("www."):
        netloc = netloc[4:]
    parts = netloc.split(".")
    if len(parts) <= 2:
        return netloc
    two_level = {"co.uk", "com.au", "co.jp", "co.nz", "com.br", "co.in",
                 "com.cn", "co.kr", "com.mx", "com.sg", "co.za"}
    if ".".join(parts[-2:]) in two_level:
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


def _name_sim(a, b):
    a, b = _core_name(a), _core_name(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def _itunes_get(url, params):
    for attempt in range(APPLE_RETRIES):
        resp = requests.get(url, params=params, timeout=15)
        if resp.status_code in (403, 429):
            time.sleep(2 ** attempt * 3)
            continue
        resp.raise_for_status()
        return resp.json().get("results", [])
    return []


def _itunes_search(term):
    return _itunes_get("https://itunes.apple.com/search",
                       {"term": term, "entity": "software",
                        "country": "us", "limit": APPLE_SEARCH_LIMIT})


def _apps_by_developer(artist_id):
    results = _itunes_get("https://itunes.apple.com/lookup",
                          {"id": artist_id, "entity": "software", "country": "us"})
    return [x for x in results if x.get("wrapperType") == "software"]


def _pick_candidate(candidates, domain, name):
    if domain:
        for c in candidates:
            if _registered_domain(c.get("sellerUrl", "")) == domain:
                return c, f"domain_match={domain}"
    best, best_sim = None, 0.0
    for c in candidates:
        sim = max(_name_sim(name, c.get("trackName", "")),
                  _name_sim(name, c.get("sellerName", "")),
                  _name_sim(name, c.get("artistName", "")))
        if sim > best_sim:
            best, best_sim = c, sim
    if best and best_sim >= 0.6:
        return best, f"name_sim={best_sim:.2f}"
    return candidates[0], "top_result_unconfirmed"


def _confidence(name, dev_name, how):
    if how.startswith("domain_match"):
        return "HIGH", how
    sim = _name_sim(name, dev_name)
    if sim >= 0.85:
        return "MEDIUM", f"{how}; dev_name_sim={sim:.2f}"
    if sim >= 0.6:
        return "LOW", f"{how}; dev_name_sim={sim:.2f}"
    return "LOW", how


def _apple_row(company, status, apple_link="", dev_id="", dev_name="", dev_url="",
               app=None, confidence="", note=""):
    app = app or {}
    return {
        "CompanyID": company["id"], "CompanyName": company["name"],
        "Website": company["website"], "PrimaryIndustryGroup": company["industry"],
        "Apple_Link_Found": apple_link, "Matched_Developer_Name": dev_name,
        "Matched_Developer_ID": dev_id, "Matched_Developer_URL": dev_url,
        "Developer_App_Name": app.get("trackName", ""),
        "Developer_App_ID": app.get("trackId", ""),
        "Developer_Bundle_ID": app.get("bundleId", ""),
        "Developer_App_URL": app.get("trackViewUrl", ""),
        "Developer_Seller_URL": app.get("sellerUrl", ""),
        "Match_Status": status, "Confidence": confidence, "Evidence_Notes": note,
    }


def _apple_process(company):
    domain = _registered_domain(company["website"])
    try:
        candidates = _itunes_search(company["name"])
        if not candidates:
            return [_apple_row(company, "NO_APP_FOUND")]
        best, how = _pick_candidate(candidates, domain, company["name"])
        artist_id = best.get("artistId")
        if not artist_id:
            return [_apple_row(company, "APP_FOUND_NO_DEVELOPER",
                               apple_link=best.get("trackViewUrl", ""))]
        apps = _apps_by_developer(artist_id) or [best]
        dev_name = apps[0].get("sellerName") or apps[0].get("artistName") or ""
        dev_url = apps[0].get("artistViewUrl") or ""
        conf, note = _confidence(company["name"], dev_name, how)
        return [_apple_row(company, "MATCHED",
                           apple_link=best.get("trackViewUrl", ""), dev_id=artist_id,
                           dev_name=dev_name, dev_url=dev_url, app=a,
                           confidence=conf, note=note) for a in apps]
    except Exception as e:
        return [_apple_row(company, "ERROR", note=str(e))]


def stage_apple():
    df = pd.read_excel(CLASSIFIED_FILE)
    if "AI_Result" not in df.columns:
        raise SystemExit("No AI_Result column — classify stage must run first.")

    all_rows, done_ids = [], set()
    if os.path.exists(APPLE_FILE):
        prev = pd.read_excel(APPLE_FILE)
        all_rows = prev.to_dict("records")
        done_ids = set(prev["CompanyID"].dropna().tolist())
        print(f"[apple] resuming — {len(done_ids)} companies already processed")

    pending = []
    for _, row in df.iterrows():
        company = {"id": row.get("CompanyID"), "name": row.get("CompanyName"),
                   "website": row.get("Website"),
                   "industry": row.get("PrimaryIndustryGroup")}
        if company["id"] in done_ids:
            continue
        status = str(row.get("AI_Result", "")).strip().upper()
        if status not in APPLE_PROCESS:
            all_rows.append(_apple_row(company, f"SKIPPED_AI_{status or 'BLANK'}"))
            done_ids.add(company["id"])
            continue
        pending.append(company)

    print(f"[apple] {len(pending)} companies to look up, {APPLE_WORKERS} parallel")
    done = 0
    with ThreadPoolExecutor(max_workers=APPLE_WORKERS) as pool:
        futures = {pool.submit(_apple_process, c): c for c in pending}
        for fut in as_completed(futures):
            all_rows.extend(fut.result())
            done += 1
            if done % APPLE_CHECKPOINT == 0:
                pd.DataFrame(all_rows).to_excel(APPLE_FILE, index=False)
                print(f"[apple]   {done}/{len(pending)} done (checkpoint saved)")

    out = pd.DataFrame(all_rows)
    out.to_excel(APPLE_FILE, index=False)
    print(f"[apple] done -> {APPLE_FILE}")
    print(out["Match_Status"].value_counts())


# =====================================================================
def main():
    refilter = "--refilter" in sys.argv

    print("=" * 50 + "\nSTEP 1 / 3  —  Filter\n" + "=" * 50)
    if refilter or not os.path.exists(PILOT_FILE):
        stage_filter()
    else:
        print(f"[filter] {PILOT_FILE} exists — skipping (use --refilter to rebuild)")

    print("=" * 50 + "\nSTEP 2 / 3  —  Classify (Claude)\n" + "=" * 50)
    stage_classify()

    print("=" * 50 + "\nSTEP 3 / 3  —  Apple lookup (iTunes API)\n" + "=" * 50)
    stage_apple()

    print("\nPipeline complete.")


if __name__ == "__main__":
    main()
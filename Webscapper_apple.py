"""
Step 3 — Apple lookup via the iTunes Search API (no DuckDuckGo, no throttling pain).

Per company (only those AI-flagged YES / UNCLEAR in step 2):
  1. iTunes SEARCH by company name -> candidate apps (each already carries artistId).
  2. Pick the best candidate: domain match on sellerUrl > name similarity > top result.
  3. Use that candidate's developer (artist) id to pull EVERY app under the developer.
  4. Save all rows with a confidence flag. Resume-safe.

Install:
    pip install pandas requests openpyxl

Note: Apple throttles the iTunes API to roughly ~20 calls/min, so this step is
inherently slower than the classifier. Keep MAX_WORKERS small; backoff handles 403s.
"""

import os
import re
import time
import pandas as pd
import requests
from urllib.parse import urlparse
from difflib import SequenceMatcher
from concurrent.futures import ThreadPoolExecutor, as_completed

INPUT_FILE = "/Users/tianyuzhou/Documents/Finance_RA/pitchbook_app_classified.xlsx"
OUTPUT_FILE = "/Users/tianyuzhou/Documents/Finance_RA/apple_developer_apps.xlsx"

PROCESS_STATUSES = {"YES", "UNCLEAR"}
MAX_WORKERS = 4            # keep small — Apple throttles the iTunes API
CHECKPOINT_EVERY = 25
SEARCH_LIMIT = 10
MAX_RETRIES = 4

NAME_NOISE = {
    "inc", "incorporated", "llc", "ltd", "limited", "corp", "corporation",
    "co", "company", "group", "holdings", "plc", "gmbh", "sa", "ag",
    "technologies", "technology", "labs", "lab", "software", "the", "app",
}


# name / domain helpers

def core_name(name):
    if not isinstance(name, str):
        return ""
    n = re.sub(r"[^a-z0-9 ]", " ", name.lower())
    return " ".join(t for t in n.split() if t and t not in NAME_NOISE).strip()


def registered_domain(url):
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


def name_similarity(a, b):
    a, b = core_name(a), core_name(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


# iTunes API (with backoff on 403/429)

def itunes_get(url, params):
    for attempt in range(MAX_RETRIES):
        resp = requests.get(url, params=params, timeout=15)
        if resp.status_code in (403, 429):
            time.sleep(2 ** attempt * 3)
            continue
        resp.raise_for_status()
        return resp.json().get("results", [])
    return []


def itunes_search(term):
    return itunes_get("https://itunes.apple.com/search",
                      {"term": term, "entity": "software", "country": "us",
                       "limit": SEARCH_LIMIT})


def get_all_apps_by_developer(artist_id):
    results = itunes_get("https://itunes.apple.com/lookup",
                         {"id": artist_id, "entity": "software", "country": "us"})
    return [x for x in results if x.get("wrapperType") == "software"]


# candidate selection + confidence

def pick_best_candidate(candidates, domain, name):
    if domain:
        for c in candidates:
            if registered_domain(c.get("sellerUrl", "")) == domain:
                return c, f"domain_match={domain}"
    best, best_sim = None, 0.0
    for c in candidates:
        sim = max(name_similarity(name, c.get("trackName", "")),
                  name_similarity(name, c.get("sellerName", "")),
                  name_similarity(name, c.get("artistName", "")))
        if sim > best_sim:
            best, best_sim = c, sim
    if best and best_sim >= 0.6:
        return best, f"name_sim={best_sim:.2f}"
    return candidates[0], "top_result_unconfirmed"


def assess_confidence(domain, name, apps, dev_name, how):
    if how.startswith("domain_match"):
        return "HIGH", how
    sim = name_similarity(name, dev_name)
    if sim >= 0.85:
        return "MEDIUM", f"{how}; dev_name_sim={sim:.2f}"
    if sim >= 0.6:
        return "LOW", f"{how}; dev_name_sim={sim:.2f}"
    return "LOW", how


def make_row(company, status, apple_link="", dev_id="", dev_name="", dev_url="",
             app=None, confidence="", note=""):
    app = app or {}
    return {
        "CompanyID": company["id"],
        "CompanyName": company["name"],
        "Website": company["website"],
        "PrimaryIndustryGroup": company["industry"],
        "Apple_Link_Found": apple_link,
        "Matched_Developer_Name": dev_name,
        "Matched_Developer_ID": dev_id,
        "Matched_Developer_URL": dev_url,
        "Developer_App_Name": app.get("trackName", ""),
        "Developer_App_ID": app.get("trackId", ""),
        "Developer_Bundle_ID": app.get("bundleId", ""),
        "Developer_App_URL": app.get("trackViewUrl", ""),
        "Developer_Seller_URL": app.get("sellerUrl", ""),
        "Match_Status": status,
        "Confidence": confidence,
        "Evidence_Notes": note,
    }


def process_company(company):
    domain = registered_domain(company["website"])
    try:
        candidates = itunes_search(company["name"])
        if not candidates:
            return [make_row(company, "NO_APP_FOUND")]

        best, how = pick_best_candidate(candidates, domain, company["name"])
        artist_id = best.get("artistId")
        if not artist_id:
            return [make_row(company, "APP_FOUND_NO_DEVELOPER",
                             apple_link=best.get("trackViewUrl", ""))]

        apps = get_all_apps_by_developer(artist_id) or [best]
        dev_name = apps[0].get("sellerName") or apps[0].get("artistName") or ""
        dev_url = apps[0].get("artistViewUrl") or ""
        confidence, note = assess_confidence(domain, company["name"], apps, dev_name, how)

        return [
            make_row(company, "MATCHED",
                     apple_link=best.get("trackViewUrl", ""), dev_id=artist_id,
                     dev_name=dev_name, dev_url=dev_url,
                     app=a, confidence=confidence, note=note)
            for a in apps
        ]
    except Exception as e:
        return [make_row(company, "ERROR", note=str(e))]


def run():
    df = pd.read_excel(INPUT_FILE)
    if "AI_Result" not in df.columns:
        raise SystemExit("No AI_Result column — run AppClassifer.py (step 2) first.")

    all_rows, done_ids = [], set()
    if os.path.exists(OUTPUT_FILE):
        prev = pd.read_excel(OUTPUT_FILE)
        all_rows = prev.to_dict("records")
        done_ids = set(prev["CompanyID"].dropna().tolist())
        print(f"[apple] resuming — {len(done_ids)} companies already processed")

    pending = []
    for _, row in df.iterrows():
        company = {
            "id": row.get("CompanyID"),
            "name": row.get("CompanyName"),
            "website": row.get("Website"),
            "industry": row.get("PrimaryIndustryGroup"),
        }
        if company["id"] in done_ids:
            continue
        status = str(row.get("AI_Result", "")).strip().upper()
        if status not in PROCESS_STATUSES:
            all_rows.append(make_row(company, f"SKIPPED_AI_{status or 'BLANK'}"))
            done_ids.add(company["id"])
            continue
        pending.append(company)

    print(f"[apple] {len(pending)} companies to look up, {MAX_WORKERS} parallel")
    done = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(process_company, c): c for c in pending}
        for fut in as_completed(futures):
            all_rows.extend(fut.result())
            done += 1
            if done % CHECKPOINT_EVERY == 0:
                pd.DataFrame(all_rows).to_excel(OUTPUT_FILE, index=False)
                print(f"[apple]   {done}/{len(pending)} done (checkpoint saved)")

    out = pd.DataFrame(all_rows)
    out.to_excel(OUTPUT_FILE, index=False)
    print(f"[apple] done -> {OUTPUT_FILE}")
    print(out["Match_Status"].value_counts())


if __name__ == "__main__":
    run()
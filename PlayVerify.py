"""
Step 4b — Verify the Google Play matches with Claude (standalone; Apple-free).

Reads the Play output (play_developer_apps.xlsx), groups rows by company:
  - HIGH confidence (package-domain match) -> auto-accepted, no API call.
  - MEDIUM / LOW matches -> sent to Claude in batches to judge.
  - Companies with no matched app -> Verified = NO.

Verdict is per company, but outputs keep one row PER APP — no app dropped.

Outputs (all Play-specific, nothing mixed with Apple):
  - play_developer_apps_verified.xlsx   (every row + Verified + Verify_Reason)
  - play_companies_with_apps_final.xlsx (every app row for VERIFIED companies)
  - play_companies_with_apps_summary.xlsx (one row per verified company)

Run after PlayScraper.py:
    python3 PlayVerify.py

Setup:
    pip install anthropic pandas openpyxl python-dotenv
    # AI_apis.env must contain ANTHROPIC_API_KEY=sk-ant-...
"""

import os
import re
import json
import time
import pandas as pd
import anthropic
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = "/Users/tianyuzhou/Documents/Finance_RA"
INPUT_FILE = f"{BASE}/play_developer_apps.xlsx"
VERIFIED_FILE = f"{BASE}/play_developer_apps_verified.xlsx"
FINAL_FILE = f"{BASE}/play_companies_with_apps_final.xlsx"
SUMMARY_FILE = f"{BASE}/play_companies_with_apps_summary.xlsx"

BATCH_SIZE = 15
MAX_WORKERS = 8
MODEL = "claude-haiku-4-5-20251001"
MAX_RETRIES = 5
APPS_SHOWN = 8  # apps listed per company in the prompt; output keeps all

SYSTEM_PROMPT = (
    "You verify whether a matched Google Play developer genuinely belongs to a "
    "given company, using the company's name, website, and industry against the "
    "developer's name and their apps (names and package IDs; package IDs are "
    "reverse-DNS, e.g. com.spinn.spinncoffee implies spinn.com). "
    "Reply with ONLY a JSON array — no markdown, no prose. One object per item: "
    '[{"index": <number>, "verdict": "YES"|"NO"}]. '
    "YES means the developer account belongs to that company."
)

_client = None


def build_message(batch):
    lines = []
    for idx, c in batch:
        apps = "; ".join(
            f"{a['name']} (package={a['package']})"
            for a in c["apps"][:APPS_SHOWN]
        )
        lines.append(
            f"[{idx}] Company: {c['name']} | Website: {c['website']} | "
            f"Industry: {c['industry']}\n"
            f"     Developer: {c['dev_name']}\n"
            f"     Apps: {apps}"
        )
    return "\n".join(lines)


def parse_verdicts(text):
    cleaned = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    out = {}
    for item in json.loads(cleaned):
        v = str(item.get("verdict", "")).strip().upper()
        if v in ("YES", "NO"):
            out[int(item["index"])] = v
    return out


def verify_batch(batch):
    msg = build_message(batch)
    for attempt in range(MAX_RETRIES):
        try:
            resp = _client.messages.create(
                model=MODEL,
                max_tokens=32 * len(batch) + 256,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": msg}],
            )
            text = "".join(b.text for b in resp.content if b.type == "text")
            return parse_verdicts(text)
        except (anthropic.RateLimitError, anthropic.APIStatusError) as e:
            if isinstance(e, anthropic.APIStatusError) and e.status_code not in (429, 500, 529):
                raise
            time.sleep(2 ** attempt * 5)
        except (json.JSONDecodeError, ValueError):
            return {}
    return {}


def build_company(cid, g):
    """Collapse a company's Play rows into one record FOR THE PROMPT ONLY."""
    first = g.iloc[0]
    matched = g[g["Match_Status"].astype(str).str.upper() == "MATCHED"]
    confs = set(matched["Confidence"].astype(str).str.upper())
    best_conf = ("HIGH" if "HIGH" in confs else
                 "MEDIUM" if "MEDIUM" in confs else
                 "LOW" if "LOW" in confs else "")
    apps = [{"name": str(r.get("Developer_App_Name", "")),
             "package": str(r.get("Developer_App_ID", ""))}
            for _, r in matched.iterrows()]
    return {
        "id": cid,
        "name": str(first.get("CompanyName", "")),
        "website": str(first.get("Website", "")),
        "industry": str(first.get("PrimaryIndustryGroup", "")),
        "dev_name": str(matched.iloc[0].get("Matched_Developer_Name", "")) if not matched.empty else "",
        "dev_id": matched.iloc[0].get("Matched_Developer_ID", "") if not matched.empty else "",
        "apps": apps,
        "app_count": len(matched),
        "best_conf": best_conf,
        "has_match": not matched.empty,
    }


def run():
    global _client
    load_dotenv("AI_apis.env")
    _client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    df = pd.read_excel(INPUT_FILE)
    companies = [build_company(cid, g) for cid, g in df.groupby("CompanyID", sort=False)]

    verdict, reason, to_ask = {}, {}, []
    for c in companies:
        if not c["has_match"]:
            verdict[c["id"]], reason[c["id"]] = "NO", "no matched app"
        elif c["best_conf"] == "HIGH":
            verdict[c["id"]], reason[c["id"]] = "YES", "auto-accepted (HIGH/package-domain match)"
        else:
            to_ask.append(c)

    print(f"[play-verify] {len(companies)} companies | "
          f"{sum(v == 'YES' for v in verdict.values())} auto-accepted | "
          f"{len(to_ask)} to check with Claude")

    indexed = list(enumerate(to_ask))
    batches = [indexed[s:s + BATCH_SIZE] for s in range(0, len(indexed), BATCH_SIZE)]
    if batches:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futures = {pool.submit(verify_batch, b): b for b in batches}
            for fut in as_completed(futures):
                res = fut.result()
                for idx, c in futures[fut]:
                    verdict[c["id"]] = res.get(idx, "NO")
                    reason[c["id"]] = f"Claude verdict ({c['best_conf'] or 'no-conf'})"

    # Broadcast the per-company verdict onto every app row — nothing dropped.
    df["Verified"] = df["CompanyID"].map(verdict).fillna("NO")
    df["Verify_Reason"] = df["CompanyID"].map(reason).fillna("")
    df.to_excel(VERIFIED_FILE, index=False)

    # FINAL: every app row belonging to a verified company (one row per app).
    status_u = df["Match_Status"].astype(str).str.upper()
    final = df[(df["Verified"] == "YES") & (status_u == "MATCHED")].copy()
    final.to_excel(FINAL_FILE, index=False)

    # SUMMARY: one row per verified company.
    summary_rows = []
    for c in companies:
        if verdict.get(c["id"]) == "YES":
            summary_rows.append({
                "CompanyID": c["id"], "CompanyName": c["name"],
                "Website": c["website"], "PrimaryIndustryGroup": c["industry"],
                "Matched_Developer_Name": c["dev_name"],
                "Matched_Developer_ID": c["dev_id"],
                "App_Count": c["app_count"], "Confidence": c["best_conf"],
                "Verify_Reason": reason.get(c["id"], ""),
            })
    pd.DataFrame(summary_rows).to_excel(SUMMARY_FILE, index=False)

    n_companies = final["CompanyID"].nunique() if not final.empty else 0
    print(f"[play-verify] {n_companies} verified companies, {len(final)} apps kept")
    print(f"[play-verify] full table   -> {VERIFIED_FILE}")
    print(f"[play-verify] every app    -> {FINAL_FILE}")
    print(f"[play-verify] per-company  -> {SUMMARY_FILE}")


if __name__ == "__main__":
    run()

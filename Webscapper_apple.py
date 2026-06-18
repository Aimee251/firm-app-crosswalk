"""
Step 3 (revised) — Web search is the entry point.

Per company (only those AI-flagged YES / UNCLEAR in step 2):
  1. Web-search the company to find ONE App Store link. Web search handles
     "company name != app name" far better than the iTunes name search.
  2. From that link, pull the App Store app id (or developer id directly).
  3. Look the app up in the iTunes API to get its DEVELOPER (artist) id.
  4. Use the developer id to pull EVERY app under that developer.
  5. Save all rows to Excel, with a confidence flag.

Install first:
    pip install ddgs pandas requests openpyxl

Note: DuckDuckGo throttles aggressively. For a 1000-company run expect it to
be slow and to occasionally return nothing (rate-limited). For production
scale, swap find_apple_listing() for a paid SERP API (Serper / SerpAPI).
"""

import os
import re
import time
import pandas as pd
import requests
from urllib.parse import urlparse
from difflib import SequenceMatcher
from ddgs import DDGS

INPUT_FILE = "/Users/tianyuzhou/Documents/Finance_RA/pitchbook_app_classified.xlsx"
OUTPUT_FILE = "/Users/tianyuzhou/Documents/Finance_RA/apple_developer_apps.xlsx"

PROCESS_STATUSES = {"YES", "UNCLEAR"}  # skip companies the AI said clearly have no app
CHECKPOINT_EVERY = 25                  # save progress every N companies
WEB_SLEEP = 3.0                        # between companies (DuckDuckGo throttles)
API_SLEEP = 1.0                        # between iTunes API calls
MAX_WEB_RESULTS = 8

NAME_NOISE = {
    "inc", "incorporated", "llc", "ltd", "limited", "corp", "corporation",
    "co", "company", "group", "holdings", "plc", "gmbh", "sa", "ag",
    "technologies", "technology", "labs", "lab", "software", "the", "app",
}


#small helpers (used for the confidence flag)

def core_name(name):
    if not isinstance(name, str):
        return ""
    n = re.sub(r"[^a-z0-9 ]", " ", name.lower())
    return " ".join(t for t in n.split() if t and t not in NAME_NOISE).strip()


def registered_domain(url):
    """Return registrable domain (e.g. 'spinn.com') from a URL/website."""
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


# web search to find the App Store link

def extract_apple_ids(url):
    """Return ('developer'|'app', id) from an apps.apple.com URL, else (None, None)."""
    if "/developer/" in url:
        m = re.search(r"id(\d+)", url)
        return ("developer", m.group(1)) if m else (None, None)
    m = re.search(r"/id(\d+)", url)  # app page: /app/<slug>/id123456789
    return ("app", m.group(1)) if m else (None, None)


def web_search(ddgs, query, retries=2):
    for attempt in range(retries + 1):
        try:
            return list(ddgs.text(query, max_results=MAX_WEB_RESULTS, region="us-en"))
        except Exception as e:
            if attempt < retries:
                time.sleep(5)
            else:
                print(f"  web search failed ({query!r}): {e}")
    return []


def find_apple_listing(company_name, company_domain, ddgs):
    """Return (apple_url, kind, apple_id) for the first App Store link found."""
    queries = [f"{company_name} app site:apps.apple.com",
               f"{company_name} app store ios"]
    if company_domain:
        queries.insert(0, f"{company_name} {company_domain} app store")

    for q in queries:
        for r in web_search(ddgs, q):
            link = r.get("href") or r.get("url") or ""
            if "apps.apple.com" in link:
                kind, apple_id = extract_apple_ids(link)
                if apple_id:
                    return link, kind, apple_id
        time.sleep(WEB_SLEEP)
    return None, None, None


# ---------- steps 2-4: iTunes API ----------

def lookup_app(app_id):
    """Look up a single app by id; return (software_item, url)."""
    url = f"https://itunes.apple.com/lookup?id={app_id}&country=us"
    resp = requests.get(url, timeout=15)
    resp.raise_for_status()
    results = resp.json().get("results", [])
    for item in results:
        if item.get("wrapperType") == "software" or item.get("kind") == "software":
            return item, url
    return (results[0] if results else None), url


def get_all_apps_by_developer(artist_id):
    """Use developer (artist) id to get every app under that developer."""
    url = f"https://itunes.apple.com/lookup?id={artist_id}&entity=software&country=us"
    resp = requests.get(url, timeout=15)
    resp.raise_for_status()
    results = resp.json().get("results", [])
    apps = [x for x in results if x.get("wrapperType") == "software"]
    return apps, url


# ---------- confidence + row builder ----------

def assess_confidence(company_domain, company_name, apps, dev_name):
    for a in apps:
        if company_domain and registered_domain(a.get("sellerUrl", "")) == company_domain:
            return "HIGH", f"domain_match={company_domain}"
    sim = name_similarity(company_name, dev_name)
    if sim >= 0.85:
        return "MEDIUM", f"dev_name_sim={sim:.2f}"
    if sim >= 0.6:
        return "LOW", f"dev_name_sim={sim:.2f}"
    return "LOW", "found via web search; no domain/name confirmation"


def make_row(company, status, apple_link="", dev_id="", dev_name="", dev_url="",
             lookup_url="", app=None, confidence="", note=""):
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
        "Developer_Lookup_URL": lookup_url,
        "Developer_App_Name": app.get("trackName", ""),
        "Developer_App_ID": app.get("trackId", ""),
        "Developer_Bundle_ID": app.get("bundleId", ""),
        "Developer_App_URL": app.get("trackViewUrl", ""),
        "Developer_Seller_URL": app.get("sellerUrl", ""),
        "Match_Status": status,
        "Confidence": confidence,
        "Evidence_Notes": note,
    }


def process_company(company, company_domain, ddgs):
    """Return the list of output rows for one company."""
    try:
        url, kind, apple_id = find_apple_listing(company["name"], company_domain, ddgs)
        if not apple_id:
            return [make_row(company, "NO_APPLE_LINK_FOUND")]

        # Get the developer (artist) id.
        if kind == "developer":
            artist_id = apple_id
        else:  # an app page -> look it up to find its developer
            app_item, _ = lookup_app(apple_id)
            time.sleep(API_SLEEP)
            if not app_item or not app_item.get("artistId"):
                return [make_row(company, "APP_FOUND_NO_DEVELOPER", apple_link=url)]
            artist_id = app_item.get("artistId")

        # Enumerate every app under that developer.
        apps, lookup_url = get_all_apps_by_developer(artist_id)
        time.sleep(API_SLEEP)

        if not apps:
            return [make_row(company, "DEVELOPER_NO_APPS",
                             apple_link=url, dev_id=artist_id, lookup_url=lookup_url)]

        dev_name = apps[0].get("sellerName") or apps[0].get("artistName") or ""
        dev_url = apps[0].get("artistViewUrl") or ""
        confidence, note = assess_confidence(company_domain, company["name"], apps, dev_name)

        return [
            make_row(company, "MATCHED", apple_link=url, dev_id=artist_id,
                     dev_name=dev_name, dev_url=dev_url, lookup_url=lookup_url,
                     app=a, confidence=confidence, note=note)
            for a in apps
        ]

    except Exception as e:
        return [make_row(company, "ERROR", note=str(e))]


if __name__ == "__main__":
    df = pd.read_excel(INPUT_FILE)
    if "AI_Result" not in df.columns:
        raise SystemExit("No AI_Result column — run AppClassifer.py (step 2) first.")

    all_rows = []
    done_ids = set()
    if os.path.exists(OUTPUT_FILE):  # resume from a previous run
        prev = pd.read_excel(OUTPUT_FILE)
        all_rows = prev.to_dict("records")
        done_ids = set(prev["CompanyID"].dropna().tolist())
        print(f"Resuming — {len(done_ids)} companies already processed.")

    total = len(df)
    processed = 0
    ddgs = DDGS()

    for index, row in df.iterrows():
        status = str(row.get("AI_Result", "")).strip().upper()
        company = {
            "id": row.get("CompanyID"),
            "name": row.get("CompanyName"),
            "website": row.get("Website"),
            "industry": row.get("PrimaryIndustryGroup"),
        }

        if company["id"] in done_ids:
            continue

        if status not in PROCESS_STATUSES:
            all_rows.append(make_row(company, f"SKIPPED_AI_{status or 'BLANK'}"))
            done_ids.add(company["id"])
            continue

        company_domain = registered_domain(company["website"])
        print(f"{index + 1}/{total}: {company['name']} (AI={status})")

        all_rows.extend(process_company(company, company_domain, ddgs))
        done_ids.add(company["id"])
        processed += 1

        if processed % CHECKPOINT_EVERY == 0:
            pd.DataFrame(all_rows).to_excel(OUTPUT_FILE, index=False)
            print(f"  ...checkpoint saved ({processed} searched this run)")

        time.sleep(WEB_SLEEP)

    pd.DataFrame(all_rows).to_excel(OUTPUT_FILE, index=False)
    print("Done.")
    print(f"Saved to: {OUTPUT_FILE}")
    print(pd.DataFrame(all_rows)["Match_Status"].value_counts())
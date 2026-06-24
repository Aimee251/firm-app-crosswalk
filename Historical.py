"""
Step 5 — Historical app detection (deterministic, evidence-based).

Goal: for each company, find apps it PREVIOUSLY had on its Apple developer page
that are NO LONGER there, and document the source + the app's ID.

How it works (no AI, no hallucination):
  1. Take the company's Apple developer page URL (from the scraper).
  2. Ask the Wayback Machine for every archived snapshot of that page over time.
  3. Parse each archived snapshot for app links (/app/<slug>/id<digits>).
     The app ID and a readable name come straight from the URL — real data.
  4. Union all app IDs ever seen historically; subtract the ones live today.
  5. Whatever remains = apps that were once on the page and aren't now.
     For each, optionally look it up live to tell DELISTED vs TRANSFERRED.

Output documents, per historical app: the app ID, name, the exact archive URL
used as evidence, the first snapshot date it appeared, and its current status.

Run after the scraper (step 3):
    python3 Historical.py
"""

import os
import re
import time
import requests
import pandas as pd

SCRAPER_OUTPUT = "/Users/tianyuzhou/Documents/Finance_RA/apple_developer_apps.xlsx"
OUTPUT_FILE = "/Users/tianyuzhou/Documents/Finance_RA/historical_apps.xlsx"

ARCHIVE_SLEEP = 1.5   # be polite to the Internet Archive
API_SLEEP = 1.0
SNAPSHOT_LIMIT = 40   # max snapshots per developer page (monthly-collapsed)

APP_LINK_RE = re.compile(r"/app/([^/?\"']+)/id(\d+)")
DEV_ID_RE = re.compile(r"id(\d+)")


def slug_to_name(slug):
    return slug.replace("-", " ").strip().title()


def developer_url_variants(developer_url):
    """Both modern and legacy Apple URL hosts, so the archive query catches more."""
    variants = {developer_url}
    if "apps.apple.com" in developer_url:
        variants.add(developer_url.replace("apps.apple.com", "itunes.apple.com"))
    if "itunes.apple.com" in developer_url:
        variants.add(developer_url.replace("itunes.apple.com", "apps.apple.com"))
    return [v for v in variants if v]


def wayback_snapshots(url, limit=SNAPSHOT_LIMIT):
    """Return archive timestamps for a URL, ~one per month."""
    try:
        r = requests.get("http://web.archive.org/cdx/search/cdx", params={
            "url": url, "output": "json", "from": "2008",
            "collapse": "timestamp:6",          # one snapshot per month
            "filter": "statuscode:200", "limit": limit,
        }, timeout=30)
        r.raise_for_status()
        data = r.json()
        return [row[1] for row in data[1:]]     # skip header row
    except Exception as e:
        print(f"    cdx error: {e}")
        return []


def apps_in_snapshot(timestamp, developer_url):
    """Fetch one archived developer page; return {app_id: name} and the source URL."""
    snap_url = f"http://web.archive.org/web/{timestamp}/{developer_url}"
    found = {}
    try:
        html = requests.get(snap_url, timeout=30).text
        for slug, app_id in APP_LINK_RE.findall(html):
            found[app_id] = slug_to_name(slug)
    except Exception as e:
        print(f"    snapshot error {timestamp}: {e}")
    return found, snap_url


def lookup_live_status(app_id, company_dev_id):
    """Is this app still on the store? If so, under whom? (DELISTED / TRANSFERRED / LIVE)."""
    try:
        r = requests.get(f"https://itunes.apple.com/lookup?id={app_id}&country=us", timeout=15)
        r.raise_for_status()
        results = r.json().get("results", [])
        if not results:
            return "DELISTED", ""
        item = results[0]
        live_dev = str(item.get("artistId", ""))
        if company_dev_id and live_dev and live_dev != str(company_dev_id):
            return "TRANSFERRED", f"now under developer id {live_dev} ({item.get('artistName')})"
        return "STILL_LIVE", ""
    except Exception as e:
        return "LOOKUP_ERROR", str(e)


if __name__ == "__main__":
    df = pd.read_excel(SCRAPER_OUTPUT)
    matched = df[df["Match_Status"].isin(["MATCHED", "WEAK_MATCH"])].copy()
    if matched.empty:
        raise SystemExit("No matched companies — run the scraper (step 3) first.")

    out_rows = []
    groups = list(matched.groupby("CompanyID"))
    total = len(groups)

    for i, (company_id, g) in enumerate(groups):
        row0 = g.iloc[0]
        company_name = row0.get("CompanyName")
        developer_url = row0.get("Matched_Developer_URL")
        dev_id = row0.get("Matched_Developer_ID")

        if not isinstance(developer_url, str) or "apple.com" not in developer_url:
            continue

        # Apps live on the page TODAY (from the scraper) — to subtract out.
        current_ids = {str(x) for x in g["Developer_App_ID"].dropna().tolist()}

        print(f"{i + 1}/{total}: {company_name}")

        # Collect every app id ever seen across archived snapshots.
        history = {}  # app_id -> (name, first_snapshot_ts, source_url)
        for variant in developer_url_variants(developer_url):
            for ts in wayback_snapshots(variant):
                snap_apps, snap_url = apps_in_snapshot(ts, variant)
                for app_id, name in snap_apps.items():
                    if app_id not in history:           # keep earliest sighting
                        history[app_id] = (name, ts, snap_url)
                time.sleep(ARCHIVE_SLEEP)

        # Historical-only = ever-seen minus currently-live.
        historical_only = {aid: v for aid, v in history.items() if aid not in current_ids}

        if not historical_only:
            continue

        for app_id, (name, first_ts, source_url) in historical_only.items():
            status, note = lookup_live_status(app_id, dev_id)
            time.sleep(API_SLEEP)
            out_rows.append({
                "CompanyID": company_id,
                "CompanyName": company_name,
                "Website": row0.get("Website"),
                "Developer_ID": dev_id,
                "Historical_App_ID": app_id,
                "Historical_App_Name": name,
                "Status": status,                       # DELISTED / TRANSFERRED / STILL_LIVE
                "Status_Note": note,
                "Evidence_Source_URL": source_url,      # the archived page proving it
                "First_Seen_Snapshot": first_ts[:8],    # YYYYMMDD
                "App_Store_ID_URL": f"https://apps.apple.com/app/id{app_id}",
            })
            print(f"    historical: {name} (id{app_id}) -> {status}")

    result = pd.DataFrame(out_rows)
    result.to_excel(OUTPUT_FILE, index=False)
    print(f"\nDone. Found {len(result)} historical apps across "
          f"{result['CompanyID'].nunique() if not result.empty else 0} companies.")
    print(f"Saved to: {OUTPUT_FILE}")

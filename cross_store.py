"""
Cross-store crosswalk — feed the APPLE results into Google Play, for free.

Idea: an iOS bundle ID and the Android package ID are usually identical
(e.g. com.spinn.spinncoffee). So instead of searching Play by company name,
we look that exact package up directly on Play. One call, exact match, and
only for companies that already have an iOS app — tiny Google footprint, so
the free Node library is plenty (no SerpApi, no cost).

Per company (seeded from the Apple matches):
  1. Take the iOS bundle IDs the Apple step found.
  2. Look each up DIRECTLY on Google Play (app mode). First hit = the Android app.
  3. From that app, get the Play developer -> the official Play developer page.
  4. Enumerate the developer's apps for a Play app count.
Output: one row per company with the official developer page on BOTH stores.

Input  : companies_with_apps_final.xlsx (Apple verified) if present,
         else apple_developer_apps.xlsx (MATCHED rows).
Output : cross_store_crosswalk.xlsx

Run:
    python3 cross_store.py

Setup:
    pip install pandas openpyxl
    npm install google-play-scraper      # in this folder, for play_bridge.mjs

Note: seeding from Apple finds companies present on BOTH stores (and Apple-only).
Android-ONLY companies (no App Store app) are out of scope here by design —
run the standalone Play pipeline on the no-iOS companies if you need those too.
"""

import os
import re
import json
import time
import subprocess
import pandas as pd
from urllib.parse import urlparse, quote
from difflib import SequenceMatcher
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = "/Users/tianyuzhou/Documents/Finance_RA"
APPLE_VERIFIED = f"{BASE}/companies_with_apps_final.xlsx"
APPLE_RAW = f"{BASE}/apple_developer_apps.xlsx"
OUTPUT_FILE = f"{BASE}/cross_store_crosswalk.xlsx"

WORKERS = 3
DEV_NUM = 200
CHECKPOINT_EVERY = 25
FALLBACK_SEARCH = True       # if no bundle id matches, try a domain-confirmed name search
RETRYABLE = {"RATE_LIMITED"}

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BRIDGE = os.path.join(SCRIPT_DIR, "play_bridge.mjs")

NAME_NOISE = {
    "inc", "incorporated", "llc", "ltd", "limited", "corp", "corporation",
    "co", "company", "group", "holdings", "plc", "gmbh", "sa", "ag",
    "technologies", "technology", "labs", "lab", "software", "the", "app",
}


class RateLimited(Exception):
    pass


# ---------- helpers ----------

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


def package_to_domain(pkg):
    if not isinstance(pkg, str):
        return ""
    parts = pkg.split(".")
    return f"{parts[1]}.{parts[0]}".lower() if len(parts) >= 2 else ""


def play_dev_page(dev_id):
    if not dev_id:
        return ""
    s = str(dev_id)
    if s.isdigit():
        return f"https://play.google.com/store/apps/dev?id={s}"
    return f"https://play.google.com/store/apps/developer?id={quote(s)}"


# ---------- Node bridge calls ----------

def _node(cmd, value, num=0):
    proc = subprocess.run(["node", BRIDGE, cmd, str(value), str(num)],
                          capture_output=True, text=True, timeout=90)
    return proc


def play_app(app_id):
    """Direct package lookup. Returns dict, or None if not on Play."""
    proc = _node("app", app_id)
    if proc.returncode == 0:
        return json.loads(proc.stdout or "null")
    err = (proc.stderr or "").lower()
    if "429" in err or "too many" in err:
        raise RateLimited(err[:200])
    return None  # 404 / not found -> no Android app under this id


def play_developer(dev_id):
    proc = _node("developer", dev_id, DEV_NUM)
    if proc.returncode != 0:
        err = (proc.stderr or "").lower()
        if "429" in err or "too many" in err:
            raise RateLimited(err[:200])
        return []
    return json.loads(proc.stdout or "[]")


def play_search(name):
    proc = _node("search", name, 20)
    if proc.returncode != 0:
        err = (proc.stderr or "").lower()
        if "429" in err or "too many" in err:
            raise RateLimited(err[:200])
        return []
    return json.loads(proc.stdout or "[]")


# ---------- build company records from the Apple file ----------

def load_apple_companies():
    src = APPLE_VERIFIED if os.path.exists(APPLE_VERIFIED) else APPLE_RAW
    print(f"[cross] reading Apple results from {src}")
    df = pd.read_excel(src)
    matched = df[df["Match_Status"].astype(str).str.upper() == "MATCHED"]

    companies = []
    for cid, g in matched.groupby("CompanyID", sort=False):
        first = g.iloc[0]
        bundle_ids = [str(b).strip() for b in g.get("Developer_Bundle_ID", pd.Series([]))
                      if isinstance(b, str) and b.strip()]
        # dedupe, keep order
        seen, ordered = set(), []
        for b in bundle_ids:
            if b not in seen:
                seen.add(b)
                ordered.append(b)
        companies.append({
            "id": cid,
            "name": str(first.get("CompanyName", "")),
            "website": str(first.get("Website", "")),
            "industry": str(first.get("PrimaryIndustryGroup", "")),
            "apple_dev_name": str(first.get("Matched_Developer_Name", "")),
            "apple_dev_id": first.get("Matched_Developer_ID", ""),
            "apple_dev_page": str(first.get("Matched_Developer_URL", "")),
            "apple_app_count": len(g),
            "bundle_ids": ordered,
        })
    return companies


# ---------- per company ----------

def crosswalk_row(company, status="DONE", has_android=False, method="",
                  play_dev_name="", play_dev_id="", play_app_count=0, note=""):
    return {
        "CompanyID": company["id"], "CompanyName": company["name"],
        "Website": company["website"], "PrimaryIndustryGroup": company["industry"],
        "has_iOS": True, "has_Android": has_android,
        "Apple_Developer_Name": company["apple_dev_name"],
        "Apple_Developer_ID": company["apple_dev_id"],
        "Apple_Developer_Page": company["apple_dev_page"],
        "Apple_App_Count": company["apple_app_count"],
        "Play_Developer_Name": play_dev_name,
        "Play_Developer_ID": play_dev_id,
        "Play_Developer_Page": play_dev_page(play_dev_id) if play_dev_id else "",
        "Play_App_Count": play_app_count,
        "Match_Method": method, "Status": status, "Notes": note,
    }


def process(company):
    try:
        play_obj, method = None, ""

        # 1) direct package-id lookups (the cheap, exact path)
        for bid in company["bundle_ids"]:
            a = play_app(bid)
            if a and a.get("developerId"):
                play_obj, method = a, f"bundle_id={bid}"
                break

        # 2) optional domain-confirmed name-search fallback
        if not play_obj and FALLBACK_SEARCH:
            domain = registered_domain(company["website"])
            for c in (play_search(company["name"]) or []):
                if domain and package_to_domain(c.get("appId", "")) == domain:
                    a = play_app(c["appId"])
                    if a and a.get("developerId"):
                        play_obj, method = a, f"name_search_domain={domain}"
                        break

        if not play_obj:
            return crosswalk_row(company, has_android=False, method="no_play_match")

        dev_id = play_obj.get("developerId")
        dev_name = play_obj.get("developer") or ""
        try:
            apps = play_developer(dev_id)
        except RateLimited:
            raise
        except Exception:
            apps = [play_obj]
        return crosswalk_row(company, has_android=True, method=method,
                             play_dev_name=dev_name, play_dev_id=dev_id,
                             play_app_count=len(apps) or 1)
    except RateLimited as e:
        return crosswalk_row(company, status="RATE_LIMITED",
                             method="RATE_LIMITED", note=str(e))
    except Exception as e:
        return crosswalk_row(company, status="ERROR", method="ERROR", note=str(e))


def run():
    if not os.path.exists(BRIDGE):
        raise SystemExit(f"Missing Node bridge: {BRIDGE}")

    companies = load_apple_companies()

    all_rows, done_ids = [], set()
    if os.path.exists(OUTPUT_FILE):
        prev = pd.read_excel(OUTPUT_FILE)
        keep = prev[~prev["Status"].astype(str).str.upper().isin(RETRYABLE)]
        all_rows = keep.to_dict("records")
        done_ids = set(keep["CompanyID"].dropna().tolist())
        print(f"[cross] resuming — {len(done_ids)} done, "
              f"{len(prev) - len(keep)} retryable rows will be redone")

    pending = [c for c in companies if c["id"] not in done_ids]
    print(f"[cross] {len(pending)} iOS companies to check on Google Play")

    done = 0
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(process, c): c for c in pending}
        for fut in as_completed(futures):
            all_rows.append(fut.result())
            done += 1
            if done % CHECKPOINT_EVERY == 0:
                pd.DataFrame(all_rows).to_excel(OUTPUT_FILE, index=False)
                print(f"[cross]   {done}/{len(pending)} done (checkpoint saved)")

    out = pd.DataFrame(all_rows)
    out.to_excel(OUTPUT_FILE, index=False)
    both = int((out["has_Android"] == True).sum())
    print(f"[cross] done -> {OUTPUT_FILE}")
    print(f"[cross] {len(out)} iOS companies | {both} also on Google Play")
    rl = out["Status"].astype(str).str.upper().isin(RETRYABLE).sum()
    if rl:
        print(f"[cross] {rl} RATE_LIMITED — rerun to retry just those.")


if __name__ == "__main__":
    run()

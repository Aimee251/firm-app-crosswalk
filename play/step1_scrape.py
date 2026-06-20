"""
Play step 1 — scrape Google Play, seeded from the Apple final Excel.

Reads companies_with_apps_final.xlsx, and for each company:
  - looks its Apple bundle IDs up DIRECTLY on Google Play (exact, free),
  - falls back to a name search if needed,
  - finds the Play developer + counts their apps,
then writes the Play columns BACK INTO the same Excel (one value per company,
broadcast onto that company's rows). No separate output file.
"""
import os
import sys
import json
import subprocess
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from play import config

WORKERS = 3
DEV_NUM = 200
RETRYABLE = {"RATE_LIMITED"}

PLAY_COLS = ["has_Android", "Play_Developer_Name", "Play_Developer_ID",
             "Play_Developer_Page", "Play_App_Count", "Play_Match_Method",
             "Play_Confidence", "Play_Status"]


def _node(cmd, value, num=0):
    return subprocess.run(["node", config.BRIDGE, cmd, str(value), str(num)],
                          capture_output=True, text=True, timeout=90)


def _app(app_id):
    p = _node("app", app_id)
    if p.returncode == 0:
        return json.loads(p.stdout or "null")
    if "429" in (p.stderr or "").lower() or "too many" in (p.stderr or "").lower():
        raise config.RateLimited((p.stderr or "")[:200])
    return None


def _developer(dev_id):
    p = _node("developer", dev_id, DEV_NUM)
    if p.returncode != 0:
        if "429" in (p.stderr or "").lower():
            raise config.RateLimited((p.stderr or "")[:200])
        return []
    return json.loads(p.stdout or "[]")


def _search(name):
    p = _node("search", name, 20)
    if p.returncode != 0:
        if "429" in (p.stderr or "").lower():
            raise config.RateLimited((p.stderr or "")[:200])
        return []
    return json.loads(p.stdout or "[]")


def _blank(status="DONE"):
    return {"has_Android": False, "Play_Developer_Name": "", "Play_Developer_ID": "",
            "Play_Developer_Page": "", "Play_App_Count": 0, "Play_Match_Method": "",
            "Play_Confidence": "", "Play_Status": status}


def _process(c):
    domain = config.registered_domain(c["website"])
    try:
        play_obj, method, conf = None, "", ""

        # 1) exact bundle-id lookups (same package on both stores)
        for bid in c["bundle_ids"]:
            a = _app(bid)
            if a and a.get("developerId"):
                play_obj, method, conf = a, f"bundle_id={bid}", "HIGH"
                break

        # 2) name search: domain match (HIGH), else best name similarity (MED/LOW)
        if not play_obj:
            cands = _search(c["name"]) or []
            for cand in cands:
                if domain and config.package_to_domain(cand.get("appId", "")) == domain:
                    a = _app(cand["appId"])
                    if a and a.get("developerId"):
                        play_obj, method, conf = a, f"name_search_domain={domain}", "HIGH"
                        break
            if not play_obj and cands:
                best, bs = None, 0.0
                for cand in cands:
                    s = max(config.name_sim(c["name"], cand.get("title", "")),
                            config.name_sim(c["name"], cand.get("developer", "")))
                    if s > bs:
                        best, bs = cand, s
                if best and bs >= 0.6:
                    a = _app(best["appId"])
                    if a and a.get("developerId"):
                        play_obj = a
                        method = f"name_sim={bs:.2f}"
                        conf = "MEDIUM" if bs >= 0.85 else "LOW"

        if not play_obj:
            r = _blank()
            r["Play_Match_Method"] = "no_play_match"
            return r

        dev_id = play_obj.get("developerId")
        dev_name = play_obj.get("developer") or ""
        try:
            apps = _developer(dev_id)
        except config.RateLimited:
            raise
        except Exception:
            apps = [play_obj]
        return {"has_Android": True, "Play_Developer_Name": dev_name,
                "Play_Developer_ID": dev_id,
                "Play_Developer_Page": config.play_dev_page(dev_id),
                "Play_App_Count": len(apps) or 1, "Play_Match_Method": method,
                "Play_Confidence": conf, "Play_Status": "DONE"}
    except config.RateLimited:
        return _blank(status="RATE_LIMITED")
    except Exception as e:
        r = _blank(status="ERROR")
        r["Play_Match_Method"] = str(e)[:120]
        return r


def run():
    if not os.path.exists(config.BRIDGE):
        raise SystemExit(f"Missing Node bridge: {config.BRIDGE}")
    if not os.path.exists(config.APPLE_FINAL):
        raise SystemExit(f"Missing Apple final: {config.APPLE_FINAL} (run the Apple side first)")

    df = pd.read_excel(config.APPLE_FINAL)
    for col in PLAY_COLS:
        if col not in df.columns:
            df[col] = "" if col != "has_Android" else False

    # build one record per company (collect bundle IDs)
    companies = {}
    for cid, g in df.groupby("CompanyID", sort=False):
        first = g.iloc[0]
        seen, ordered = set(), []
        for b in g.get("Developer_Bundle_ID", pd.Series([])):
            if isinstance(b, str) and b.strip() and b not in seen:
                seen.add(b)
                ordered.append(b.strip())
        companies[cid] = {"id": cid, "name": str(first.get("CompanyName", "")),
                          "website": str(first.get("Website", "")),
                          "bundle_ids": ordered,
                          "status": str(first.get("Play_Status", "")).strip().upper()}

    pending = [c for c in companies.values()
               if c["status"] not in {"DONE"} or c["status"] in RETRYABLE]
    print(f"[play1] {len(pending)} companies to check on Google Play")

    results = {}
    done = 0
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(_process, c): c for c in pending}
        for fut in as_completed(futures):
            c = futures[fut]
            results[c["id"]] = fut.result()
            done += 1
            if done % 25 == 0:
                print(f"[play1]   {done}/{len(pending)}")

    # write each Play column back into the same file, broadcast per company
    for col in PLAY_COLS:
        mapping = {cid: r[col] for cid, r in results.items()}
        df[col] = df.apply(lambda row: mapping.get(row["CompanyID"], row[col]), axis=1)

    df.to_excel(config.APPLE_FINAL, index=False)
    both = sum(1 for r in results.values() if r["has_Android"])
    print(f"[play1] done -> wrote Play columns into {config.APPLE_FINAL}")
    print(f"[play1] {len(results)} companies checked | {both} found on Google Play")


if __name__ == "__main__":
    run()

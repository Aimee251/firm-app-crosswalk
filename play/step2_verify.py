"""
Play step 2 — verify the Google Play matches with Claude, same method as Apple.

  - HIGH confidence (exact bundle-id or domain match) -> auto-accepted, no API call.
  - MEDIUM / LOW (name-similarity matches) -> Claude judges YES/NO.
  - No Play app -> Play_Verified = NO.

Adds Play_Verified + Play_Verify_Reason columns into the SAME Excel.
"""
import os
import sys
import re
import json
import time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd
import anthropic
from concurrent.futures import ThreadPoolExecutor, as_completed
from play import config

BATCH_SIZE = 15
WORKERS = 8
RETRIES = 5
SYSTEM = (
    "You verify whether a matched Google Play developer genuinely belongs to a "
    "given company, using the company's name, website, and industry against the "
    "developer's name and how the match was made. "
    "Reply with ONLY a JSON array — no markdown, no prose. One object per item: "
    '[{"index": <number>, "verdict": "YES"|"NO"}]. '
    "YES means the developer account belongs to that company."
)


def _message(batch):
    lines = []
    for idx, c in batch:
        lines.append(
            f"[{idx}] Company: {c['name']} | Website: {c['website']} | "
            f"Industry: {c['industry']}\n"
            f"     Play developer: {c['dev_name']} | match: {c['method']}"
        )
    return "\n".join(lines)


def _parse(text):
    cleaned = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    out = {}
    for item in json.loads(cleaned):
        v = str(item.get("verdict", "")).strip().upper()
        if v in ("YES", "NO"):
            out[int(item["index"])] = v
    return out


def _batch(batch, client):
    for attempt in range(RETRIES):
        try:
            resp = client.messages.create(
                model=config.MODEL, max_tokens=32 * len(batch) + 256,
                system=SYSTEM, messages=[{"role": "user", "content": _message(batch)}],
            )
            text = "".join(b.text for b in resp.content if b.type == "text")
            return _parse(text)
        except (anthropic.RateLimitError, anthropic.APIStatusError) as e:
            if isinstance(e, anthropic.APIStatusError) and e.status_code not in (429, 500, 529):
                raise
            time.sleep(2 ** attempt * 5)
        except (json.JSONDecodeError, ValueError):
            return {}
    return {}


def run():
    client = config.get_client()
    df = pd.read_excel(config.APPLE_FINAL)
    if "has_Android" not in df.columns:
        raise SystemExit("No Play columns — run play step 1 (scrape) first.")

    verdict, reason, to_ask = {}, {}, []
    for cid, g in df.groupby("CompanyID", sort=False):
        first = g.iloc[0]
        has_android = bool(first.get("has_Android"))
        conf = str(first.get("Play_Confidence", "")).strip().upper()
        if not has_android:
            verdict[cid], reason[cid] = "NO", "no play app"
        elif conf == "HIGH":
            verdict[cid], reason[cid] = "YES", "auto-accepted (exact/domain match)"
        else:
            to_ask.append({"id": cid, "name": str(first.get("CompanyName", "")),
                           "website": str(first.get("Website", "")),
                           "industry": str(first.get("PrimaryIndustryGroup", "")),
                           "dev_name": str(first.get("Play_Developer_Name", "")),
                           "method": str(first.get("Play_Match_Method", ""))})

    print(f"[play2] {sum(v == 'YES' for v in verdict.values())} auto-accepted | "
          f"{len(to_ask)} to check with Claude")

    indexed = list(enumerate(to_ask))
    batches = [indexed[s:s + BATCH_SIZE] for s in range(0, len(indexed), BATCH_SIZE)]
    if batches:
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            futures = {pool.submit(_batch, b, client): b for b in batches}
            for fut in as_completed(futures):
                res = fut.result()
                for idx, c in futures[fut]:
                    verdict[c["id"]] = res.get(idx, "NO")
                    reason[c["id"]] = "Claude verdict"

    df["Play_Verified"] = df["CompanyID"].map(verdict).fillna("NO")
    df["Play_Verify_Reason"] = df["CompanyID"].map(reason).fillna("")
    df.to_excel(config.APPLE_FINAL, index=False)

    yes = sum(1 for v in verdict.values() if v == "YES")
    print(f"[play2] done -> {config.APPLE_FINAL}")
    print(f"[play2] {yes} companies verified on Google Play")


if __name__ == "__main__":
    run()

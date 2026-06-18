"""
Step 2 — AI pre-filter: does this company currently have mobile app(s)?
Writes an AI_Result column (YES / NO / UNCLEAR). Step 3 only searches
Apple for YES/UNCLEAR companies.
"""

import os
import time
import pandas as pd
from google import genai
from dotenv import load_dotenv

INPUT_FILE = "/Users/tianyuzhou/Documents/Finance_RA/pitchbook_app_pilot.xlsx"
OUTPUT_FILE = "/Users/tianyuzhou/Documents/Finance_RA/pitchbook_app_classified.xlsx"

CHECKPOINT_EVERY = 25
SLEEP_SECONDS = 1.0
MODEL = "gemini-2.5-flash"

load_dotenv("AI_apis.env")  # your env file is named AI_apis.env, not .env
client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

VALID_LABELS = ("YES", "NO", "UNCLEAR")


def classify_company(company, description, website, industry_group, industry_code):
    prompt = f"""
Determine whether this company currently has its own mobile app(s),
based on the description, website, and industry.

Company: {company}
Website: {website}
Description: {description}
PrimaryIndustryGroup: {industry_group}
PrimaryIndustryCode: {industry_code}

Answer with ONLY ONE word: YES, NO, or UNCLEAR.
"""
    try:
        response = client.models.generate_content(model=MODEL, contents=prompt)
        answer = (response.text or "").strip().upper()
    except Exception as e:
        return f"ERROR: {e}"

    for label in VALID_LABELS:
        if label in answer:
            return label
    return "UNCLEAR"


if __name__ == "__main__":
    if os.path.exists(OUTPUT_FILE):
        df = pd.read_excel(OUTPUT_FILE)
        print(f"Resuming from {OUTPUT_FILE}")
    else:
        df = pd.read_excel(INPUT_FILE)

    if "AI_Result" not in df.columns:
        df["AI_Result"] = ""

    total = len(df)
    for i, row in df.iterrows():
        current = str(df.at[i, "AI_Result"]).strip().upper()
        if current in VALID_LABELS:
            continue

        result = classify_company(
            row.get("CompanyName"),
            row.get("Description"),
            row.get("Website"),
            row.get("PrimaryIndustryGroup"),
            row.get("PrimaryIndustryCode"),
        )
        df.at[i, "AI_Result"] = result
        print(f"{i + 1}/{total}: {row.get('CompanyName')} -> {result}")

        if (i + 1) % CHECKPOINT_EVERY == 0:
            df.to_excel(OUTPUT_FILE, index=False)

        time.sleep(SLEEP_SECONDS)

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Done. Saved to: {OUTPUT_FILE}")
    print(df["AI_Result"].value_counts())
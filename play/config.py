"""Shared config + helpers for the Google Play steps."""
import os
import re
from urllib.parse import urlparse, quote
from difflib import SequenceMatcher
from dotenv import load_dotenv
import anthropic

BASE = "/Users/tianyuzhou/Documents/Finance_RA"
# The Apple verified final — we read it AND write the Play answers back into it.
APPLE_FINAL = f"{BASE}/companies_with_apps_final.xlsx"
MODEL = "claude-haiku-4-5-20251001"

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(THIS_DIR)
BRIDGE = os.path.join(PROJECT_DIR, "play_bridge.mjs")
ENV_FILE = os.path.join(PROJECT_DIR, "AI_apis.env")

NAME_NOISE = {
    "inc", "incorporated", "llc", "ltd", "limited", "corp", "corporation",
    "co", "company", "group", "holdings", "plc", "gmbh", "sa", "ag",
    "technologies", "technology", "labs", "lab", "software", "the", "app",
}


class RateLimited(Exception):
    pass


_client = None


def get_client():
    global _client
    if _client is None:
        load_dotenv(ENV_FILE)
        _client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    return _client


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


def name_sim(a, b):
    a, b = core_name(a), core_name(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


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

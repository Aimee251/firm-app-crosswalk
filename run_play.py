"""
ONE trigger for the whole crosswalk.

    python run_play.py

- Apple side is LEFT AS-IS: if the Apple final Excel doesn't exist yet, this
  runs your existing run_pipeline.py + Verify.py to create it. If it already
  exists, Apple is skipped entirely (untouched).
- Then the Google Play steps run, writing their answers INTO that same Excel:
    play/step1_scrape.py  — find Play apps + developer page (free Node)
    play/step2_verify.py  — Claude verifies the Play matches

Final file (both stores in one): companies_with_apps_final.xlsx
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from play import config, step1_scrape, step2_verify


def main():
    print(f"### crosswalk | final file: {config.APPLE_FINAL} ###\n")

    if not os.path.exists(config.APPLE_FINAL):
        print("Apple final not found — running the Apple side first (left as-is)…")
        import run_pipeline          # your existing Apple pipeline
        import Verify                # your existing Apple verify
        run_pipeline.main()          # filter + classify + apple
        Verify.run()                 # -> companies_with_apps_final.xlsx
    else:
        print("Apple final already exists — leaving Apple untouched.")

    print("\n" + "=" * 50 + "\nGOOGLE PLAY 1 / 2 — scrape\n" + "=" * 50)
    step1_scrape.run()
    print("\n" + "=" * 50 + "\nGOOGLE PLAY 2 / 2 — verify (Claude)\n" + "=" * 50)
    step2_verify.run()

    print(f"\nDone. Both stores are in: {config.APPLE_FINAL}")


if __name__ == "__main__":
    main()

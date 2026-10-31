#!/usr/bin/env python3
"""Fetch FDA eCFR Title 21 parts and store the raw section text.

Tracker tasks: P4-01 (eCFR API access check), P4-02 (eCFR poller)
Prompts: ecfr_title, ecfr_parts (comma-separated), output_directory,
         http_proxy (blank = none)
Prerequisites: Python 3.12, network route to ecfr.gov (direct or via proxy).
  The eCFR API is free and needs no key. Docs: ecfr.gov/developers.
Output: one JSON file per part: <output>/ecfr_title-<title>_part-<part>.json
  with section numbers and raw text.
"""
import json
import os
import urllib.request

ECFR_BASE_URL = "https://www.ecfr.gov/api/renderer/v1/content/enhanced/current/title-{}"


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_optional(label, default=""):
    return input(f"{label} [{default}]: ").strip() or default


def fetch_json(url, proxy):
    opener = urllib.request.build_opener(
        urllib.request.ProxyHandler({"http": proxy, "https": proxy}) if proxy else None)
    request = urllib.request.Request(url, headers={"Accept": "application/json"})
    with opener.open(request, timeout=60) as response:
        return json.loads(response.read().decode("utf-8"))


def main():
    print("=== Step 1: eCFR settings ===")
    ecfr_title = prompt("eCFR title (example — replace with your value)", "21")
    ecfr_parts = prompt("Parts, comma-separated (example — replace with your value)", "210,211,11")
    output_directory = prompt("Output directory (example — replace with your value)", "/data/staging/ecfr")
    http_proxy = prompt_optional("HTTP proxy (blank = direct)")

    print("=== Step 2: Connectivity check (P4-01) ===")
    titles = fetch_json("https://www.ecfr.gov/api/versioner/v1/titles.json", http_proxy)
    print(f"eCFR API OK ({len(titles.get('titles', []))} titles available).")

    print("=== Step 3: Fetching parts ===")
    os.makedirs(output_directory, exist_ok=True)
    for part in [item.strip() for item in ecfr_parts.split(",") if item.strip()]:
        url = f"{ECFR_BASE_URL.format(ecfr_title)}?part={part}"
        print(f"Fetching part {part} ...")
        try:
            data = fetch_json(url, http_proxy)
        except Exception as error:
            print(f"  WARNING: part {part} failed: {error}")
            continue
        out_path = os.path.join(output_directory, f"ecfr_title-{ecfr_title}_part-{part}.json")
        with open(out_path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2)
        print(f"  Saved {out_path}")
    print("Done. Next: 04-parse-ecfr-clauses.py.")


if __name__ == "__main__":
    main()

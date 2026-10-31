#!/usr/bin/env python3
"""Download ICH guideline PDFs.

Tracker task: P4-05 (ICH fetcher)
Prompts: ich_guidelines (comma-separated, e.g. Q7,Q8,Q9,Q10,Q12),
         output_directory, http_proxy (blank = none),
         pdf_url (per guideline, blank = skip)
Prerequisites: Python 3.12, network route to ich.org.
  Like the EU GMP fetcher, this asks for the current PDF URL per guideline
  because ich.org reorganises links; paste them when prompted.
Output: PDFs in <output_directory>, plus a manifest JSON.
"""
import json
import os
import urllib.request

ICH_LANDING = "https://www.ich.org/page/multidisciplinary-guidelines"


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_optional(label, default=""):
    return input(f"{label} [{default}]: ").strip() or default


def main():
    print("=== Step 1: ICH selection ===")
    print(f"Source landing page (verify current links there): {ICH_LANDING}")
    ich_guidelines = prompt("Guidelines to fetch, comma-separated (example — replace with your value)", "Q7,Q8,Q9,Q10,Q12")
    output_directory = prompt("Output directory (example — replace with your value)", "/data/staging/ich")
    http_proxy = prompt_optional("HTTP proxy (blank = direct)")
    opener = urllib.request.build_opener(
        urllib.request.ProxyHandler({"http": http_proxy, "https": http_proxy}) if http_proxy else None)

    os.makedirs(output_directory, exist_ok=True)
    manifest = []
    print("=== Step 2: Downloading (paste each current PDF URL when asked) ===")
    for guideline in [entry.strip() for entry in ich_guidelines.split(",") if entry.strip()]:
        pdf_url = prompt(f"PDF URL for ICH {guideline} (blank = skip)")
        if not pdf_url:
            print(f"  Skipped ICH {guideline}.")
            continue
        out_path = os.path.join(output_directory, f"ich-{guideline}.pdf")
        try:
            with opener.open(pdf_url, timeout=120) as response, open(out_path, "wb") as handle:
                handle.write(response.read())
            manifest.append({"guideline": guideline, "url": pdf_url, "file": out_path})
            print(f"  Saved {out_path}")
        except Exception as error:
            print(f"  WARNING: ICH {guideline} failed: {error}")
    with open(os.path.join(output_directory, "manifest.json"), "w", encoding="utf-8") as handle:
        json.dump(manifest, handle, indent=2)
    print(f"Done. {len(manifest)} PDFs in {output_directory}. Next: 04-split-standard-pdfs.py.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Download EU GMP (EudraLex Volume 4) chapter/annex PDFs.

Tracker task: P4-04 (EU GMP fetcher)
Prompts: eugmp_chapters (comma-separated), eugmp_annexes (comma-separated),
         output_directory, http_proxy (blank = none)
Prerequisites: Python 3.12, network route to health.ec.europa.eu.
  The script asks for the document URLs because the EC site reorganises
  links; paste the current PDF URLs when prompted. If a URL fails, the
  script logs it and continues with the rest.
Output: PDFs in <output_directory>, plus a manifest JSON of what was fetched.
"""
import json
import os
import urllib.request

EUDRALEX_LANDING = "https://health.ec.europa.eu/medicinal-products/eudralex/eudralex-volume-4_en"


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_optional(label, default=""):
    return input(f"{label} [{default}]: ").strip() or default


def main():
    print("=== Step 1: EU GMP selection ===")
    print(f"Source landing page (verify current links there): {EUDRALEX_LANDING}")
    eugmp_chapters = prompt("Chapters to fetch, comma-separated", "1,2,3,4,5,6,7,8,9")
    eugmp_annexes = prompt("Annexes to fetch, comma-separated", "1,11,15")
    output_directory = prompt("Output directory", "/data/staging/eugmp")
    http_proxy = prompt_optional("HTTP proxy (blank = direct)")
    opener = urllib.request.build_opener(
        urllib.request.ProxyHandler({"http": http_proxy, "https": http_proxy}) if http_proxy else None)

    os.makedirs(output_directory, exist_ok=True)
    manifest = []
    print("=== Step 2: Downloading (paste each current PDF URL when asked) ===")
    for kind, items in (("chapter", eugmp_chapters.split(",")), ("annex", eugmp_annexes.split(","))):
        for item in [entry.strip() for entry in items if entry.strip()]:
            pdf_url = prompt(f"PDF URL for {kind} {item} (blank = skip)")
            if not pdf_url:
                print(f"  Skipped {kind} {item}.")
                continue
            out_path = os.path.join(output_directory, f"eugmp-{kind}-{item}.pdf")
            try:
                with opener.open(pdf_url, timeout=120) as response, open(out_path, "wb") as handle:
                    handle.write(response.read())
                manifest.append({"kind": kind, "number": item, "url": pdf_url, "file": out_path})
                print(f"  Saved {out_path}")
            except Exception as error:
                print(f"  WARNING: {kind} {item} failed: {error}")
    with open(os.path.join(output_directory, "manifest.json"), "w", encoding="utf-8") as handle:
        json.dump(manifest, handle, indent=2)
    print(f"Done. {len(manifest)} PDFs in {output_directory}. Next: 04-split-standard-pdfs.py.")


if __name__ == "__main__":
    main()

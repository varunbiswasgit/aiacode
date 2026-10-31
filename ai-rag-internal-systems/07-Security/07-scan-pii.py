#!/usr/bin/env python3
"""Scan ingested chunks for PII patterns and quarantine flagged documents.

Tracker task: P7-04 (PII scan on ingestion)
Prompts: input_jsonl_path, output_clean_path, output_quarantine_path
Prerequisites: Python 3.12, chunk JSONL from phase 03.
Patterns: US SSN, credit-card-like sequences, email addresses.
  Flagged chunks go to the quarantine file for human review; everything
  else goes to the clean file. This is a pattern scan, not a guarantee -
  review quarantined docs before releasing them.
Optional: if presidio-analyzer is installed it is NOT used here, to keep
  this script dependency-free; swap the pattern list for Presidio later.
"""
import json
import re

PATTERNS = {
    "us_ssn": re.compile(r"\b\d{3}-\d{2}-\d{4}\b"),
    "credit_card_like": re.compile(r"\b(?:\d[ -]?){13,19}\b"),
    "email": re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"),
}


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def main():
    print("=== Step 1: Scan settings ===")
    input_jsonl_path = prompt("Input chunk JSONL", "/data/staging/chunks-pilot.jsonl")
    output_clean_path = prompt("Clean output JSONL", "/data/staging/chunks-clean.jsonl")
    output_quarantine_path = prompt("Quarantine output JSONL", "/data/staging/chunks-quarantine.jsonl")

    print("=== Step 2: Scanning ===")
    clean = 0
    quarantined = 0
    quarantined_docs = set()
    with open(input_jsonl_path, encoding="utf-8") as source, \
            open(output_clean_path, "w", encoding="utf-8") as clean_out, \
            open(output_quarantine_path, "w", encoding="utf-8") as quarantine_out:
        for line in source:
            record = json.loads(line)
            text = record.get("text", "")
            hits = [name for name, pattern in PATTERNS.items() if pattern.search(text)]
            if hits:
                record["pii_flags"] = hits
                quarantine_out.write(json.dumps(record) + "\n")
                quarantined += 1
                quarantined_docs.add(record.get("doc_id"))
            else:
                clean_out.write(json.dumps(record) + "\n")
                clean += 1
    print(f"Done. {clean} clean chunks, {quarantined} quarantined chunks.")
    if quarantined_docs:
        print("Quarantined documents for human review:")
        for doc_id in sorted(quarantined_docs):
            print(f"  - {doc_id}")
        print("Only load the CLEAN file into the vector store until review is done.")


if __name__ == "__main__":
    main()

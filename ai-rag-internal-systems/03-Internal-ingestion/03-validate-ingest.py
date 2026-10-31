#!/usr/bin/env python3
"""Spot-check ingest quality: show random chunks, ask a human to confirm them.

Tracker task: P3-10 (validate ingest quality)
Prompts: input_jsonl_path, sample_size
Prerequisites: chunk JSONL from 03-ingest-corpus.py or 03-parse-documents.py.
For each sampled chunk the reviewer answers y (good) or n (bad) and the
script reports a pass rate. Fix failures, then re-run the failed docs only.
"""
import json
import random


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_int(label, default):
    while True:
        raw = input(f"{label} [{default}]: ").strip() or str(default)
        if raw.isdigit() and int(raw) > 0:
            return int(raw)
        print("Enter a positive whole number.")


def main():
    print("=== Step 1: Sample settings ===")
    input_jsonl_path = prompt("Input chunk JSONL (example — replace with your value)", "/data/staging/chunks-pilot.jsonl")
    sample_size = prompt_int("How many random chunks to review (example — replace with your value)", 20)

    with open(input_jsonl_path, encoding="utf-8") as handle:
        chunks = [json.loads(line) for line in handle]
    if not chunks:
        print("ERROR: no chunks found in the input file.")
        raise SystemExit(1)
    sample = random.sample(chunks, min(sample_size, len(chunks)))

    print(f"=== Step 2: Reviewing {len(sample)} chunks ===")
    good = 0
    bad_docs = []
    for number, chunk in enumerate(sample, 1):
        print(f"\n--- Chunk {number}/{len(sample)} | doc: {chunk.get('doc_id')} | index: {chunk.get('chunk_index')} ---")
        print(chunk.get("text", "")[:800])
        answer = input("Does this chunk look correct vs the source doc? (y/n): ").strip().lower()
        if answer == "y":
            good += 1
        else:
            bad_docs.append(chunk.get("doc_id"))
    print(f"\nDone. Pass rate: {good}/{len(sample)} ({100 * good // len(sample)}%).")
    if bad_docs:
        print("Docs needing parser fixes / re-ingest:")
        for doc_id in sorted(set(bad_docs)):
            print(f"  - {doc_id}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Build the golden Q&A evaluation set interactively, one item at a time.

Tracker task: P9-01 (golden Q&A set)
Prompts: output_golden_path, then per item: question, expected_answer,
         expected_doc_ids (comma-separated). A blank question ends entry.
Prerequisites: none. The golden set should cover the pilot corpus and the
  standards (aim for 50-100 items before running 09-run-evaluation.py).
Output: golden.json - a list of {question, expected_answer, expected_doc_ids}.
  Version this file in the repo; never edit it without bumping its version note.
"""
import json


def prompt(label, default=""):
    return input(f"{label} [{default}]: ").strip() or default


def main():
    print("=== Step 1: Output ===")
    output_golden_path = prompt("Golden set output path (example — replace with your value)", "/opt/airag/eval/golden.json")
    existing = []
    try:
        with open(output_golden_path, encoding="utf-8") as handle:
            existing = json.load(handle)
        print(f"Loaded {len(existing)} existing items (new items will be appended).")
    except (FileNotFoundError, json.JSONDecodeError):
        print("Starting a new golden set.")

    print("=== Step 2: Add items (blank question = finished) ===")
    items = list(existing)
    while True:
        question = input("\nQuestion (blank to finish): ").strip()
        if not question:
            break
        expected_answer = input("Expected answer (short): ").strip()
        expected_doc_ids = input("Expected source doc IDs, comma-separated: ").strip()
        items.append({
            "question": question,
            "expected_answer": expected_answer,
            "expected_doc_ids": [entry.strip() for entry in expected_doc_ids.split(",")
                                 if entry.strip()],
        })
        print(f"  Added ({len(items)} total).")

    with open(output_golden_path, "w", encoding="utf-8") as handle:
        json.dump(items, handle, indent=2)
    print(f"\nDone. {len(items)} golden items -> {output_golden_path}")
    print("Aim for 50-100 items covering both the pilot corpus and the standards.")


if __name__ == "__main__":
    main()

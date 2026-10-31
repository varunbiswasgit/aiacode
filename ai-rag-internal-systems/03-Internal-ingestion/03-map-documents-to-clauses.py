#!/usr/bin/env python3
"""Map internal documents to the standard clauses they implement.

Tracker task: P3-11 (doc-to-clause mapping pass)
Prompts: chunks_jsonl_path, clauses_jsonl_path, output_mappings_path,
         keyword_overlap_threshold
Prerequisites: chunk JSONL (internal docs) and clause JSONL with fields
  standard, chapter, clause_no, text (from phase 04 scripts).
Method: (1) explicit citations, e.g. '21 CFR 211.100', 'ICH Q7 5.2',
  'EU GMP Annex 11'; (2) keyword overlap above the threshold.
Output: mappings JSONL with status 'confirmed' (citation) or 'needs_review'
  (keyword match only). Review 'needs_review' rows in the app review queue.
"""
import json
import os
import re
import string

CITATION_PATTERNS = [
    ("FDA eCFR", re.compile(r"21\s*CFR\s*(\d+)\.(\d+)", re.IGNORECASE)),
    ("ICH", re.compile(r"\bICH\s*(Q\d+)\s*(?:section\s*)?([\d.]+)?", re.IGNORECASE)),
    ("EU GMP", re.compile(r"\b(?:Annex|Chapter)\s*(\d+)\b", re.IGNORECASE)),
]


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_float(label, default):
    while True:
        raw = input(f"{label} [{default}]: ").strip() or str(default)
        try:
            value = float(raw)
            if 0 < value < 1:
                return value
        except ValueError:
            pass
        print("Enter a decimal between 0 and 1.")


def tokenize(text):
    translator = str.maketrans("", "", string.punctuation)
    return set(text.lower().translate(translator).split())


def main():
    print("=== Step 1: Mapping settings ===")
    chunks_jsonl_path = prompt("Internal chunks JSONL (example — replace with your value)", "/data/staging/chunks-pilot.jsonl")
    clauses_jsonl_path = prompt("Clauses JSONL (example — replace with your value)", "/data/staging/clauses.jsonl")
    output_mappings_path = prompt("Output mappings JSONL (example — replace with your value)", "/data/staging/mappings.jsonl")
    keyword_overlap_threshold = prompt_float("Keyword overlap threshold (0-1) (example — replace with your value)", 0.25)

    with open(clauses_jsonl_path, encoding="utf-8") as handle:
        clauses = [json.loads(line) for line in handle]
    clause_tokens = [(clause, tokenize(clause.get("text", ""))) for clause in clauses]
    print(f"Loaded {len(clauses)} clauses.")

    print("=== Step 2: Mapping ===")
    confirmed = 0
    needs_review = 0
    with open(chunks_jsonl_path, encoding="utf-8") as source, \
            open(output_mappings_path, "w", encoding="utf-8") as output:
        for line in source:
            chunk = json.loads(line)
            text = chunk.get("text", "")
            mapped = {}
            for standard, pattern in CITATION_PATTERNS:
                for match in pattern.finditer(text):
                    mapped[f"{standard}:{match.group(0)}"] = "confirmed"
            chunk_tokens = tokenize(text)
            for clause, tokens in clause_tokens:
                if not tokens or not chunk_tokens:
                    continue
                overlap = len(tokens & chunk_tokens) / len(tokens)
                key = f"{clause.get('standard')}:{clause.get('clause_no')}"
                if overlap >= keyword_overlap_threshold and key not in mapped:
                    mapped[key] = "needs_review"
            for clause_key, status in mapped.items():
                output.write(json.dumps({
                    "doc_id": chunk.get("doc_id"), "chunk_index": chunk.get("chunk_index"),
                    "clause_key": clause_key, "status": status,
                }) + "\n")
                if status == "confirmed":
                    confirmed += 1
                else:
                    needs_review += 1
    os.chmod(output_mappings_path, 0o600)
    print(f"Done. {confirmed} citation-confirmed, {needs_review} keyword-only (needs_review).")
    print("Review the 'needs_review' rows in the app review queue before go-live.")


if __name__ == "__main__":
    main()

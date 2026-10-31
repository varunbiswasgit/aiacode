#!/usr/bin/env python3
"""Enrich chunk JSONL with document metadata (title, owner, version, dates).

Tracker task: P3-06 (metadata extractor)
Prompts: input_jsonl_path, output_jsonl_path, doc_id_pattern (regex with a
         named group 'docid'), default_owner, default_department
Prerequisites: chunk JSONL from 03-parse-documents.py.
Output: same JSONL plus metadata fields on every chunk.
"""
import json
import re


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def main():
    print("=== Step 1: Metadata settings ===")
    input_jsonl_path = prompt("Input chunk JSONL", "/data/staging/chunks.jsonl")
    output_jsonl_path = prompt("Output enriched JSONL", "/data/staging/chunks-meta.jsonl")
    doc_id_pattern = prompt("Doc-ID regex (named group 'docid')",
                            r"(?P<docid>[A-Z]{2,}-\d{3,})")
    default_owner = prompt("Default document owner", "QA")
    default_department = prompt("Default department", "Quality")
    try:
        pattern = re.compile(doc_id_pattern)
    except re.error as error:
        print(f"ERROR: invalid regex: {error}")
        raise SystemExit(1)

    print("=== Step 2: Enriching ===")
    count = 0
    with open(input_jsonl_path, encoding="utf-8") as source, \
            open(output_jsonl_path, "w", encoding="utf-8") as output:
        for line in source:
            record = json.loads(line)
            match = pattern.search(record.get("source_path", ""))
            record["metadata"] = {
                "doc_id_from_filename": match.group("docid") if match else None,
                "title": record["doc_id"].replace("_", " ").replace("/", " - "),
                "owner": default_owner,
                "department": default_department,
                "version": None,          # fill from your DMS export if available
                "effective_date": None,   # fill from your DMS export if available
            }
            output.write(json.dumps(record) + "\n")
            count += 1
    print(f"Done. {count} chunks enriched -> {output_jsonl_path}")
    print("NOTE: version/effective_date are null unless your DMS export provides them;")
    print("      join them in by doc_id before loading into the vector store.")


if __name__ == "__main__":
    main()

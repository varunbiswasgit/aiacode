#!/usr/bin/env python3
"""Split eCFR part JSON into individual clause records (clause JSONL).

Tracker task: P4-03 (clause parser for eCFR)
Prompts: ecfr_json_directory, output_clauses_path
Prerequisites: JSON files from 04-fetch-ecfr.py.
Output JSONL rows: standard, chapter, clause_no, clause_path, text,
  version, source_url.
  clause_no is the LEAF label only (e.g. '211.100', '211.22(a)'), matching
  the form citations use, so the mapping join in
  03-map-documents-to-clauses.py and the registry key in
  04-create-clause-registry.sql match it. clause_path keeps the full
  ancestor path (e.g. '211/211.100') for traceability; it is never used
  as a join key.
  version = the fetch date (YYYY-MM-DD); re-running a fetch creates a new
  version, which 04-detect-standard-changes.py diffs against the previous one.
"""
import datetime
import json
import os


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def walk_sections(node, part, rows, path=""):
    """Recursively collect leaf sections (those with text) from eCFR JSON."""
    if isinstance(node, dict):
        label = str(node.get("label", "") or "").strip()
        section_path = f"{path}/{label}" if path else label
        # Leaf label only (e.g. '211.100', not the full '211/211.100'
        # path). Fall back to the last path segment when the node
        # itself carries no label, so clause_no is never empty while
        # a path exists.
        clause_no = label or section_path.rsplit("/", 1)[-1]
        text = node.get("text") or node.get("content")
        if isinstance(text, str) and text.strip():
            rows.append({
                "standard": "FDA eCFR",
                "chapter": f"Title 21 Part {part}",
                "clause_no": clause_no,
                "clause_path": section_path,
                "text": text.strip(),
                "version": datetime.date.today().isoformat(),
                "source_url": "https://www.ecfr.gov/current/title-21",
            })
        for child in node.get("children", []) or node.get("sections", []) or []:
            walk_sections(child, part, rows, section_path)
    elif isinstance(node, list):
        for child in node:
            walk_sections(child, part, rows, path)


def main():
    print("=== Step 1: Parser settings ===")
    ecfr_json_directory = prompt("Directory with eCFR part JSON", "/data/staging/ecfr")
    output_clauses_path = prompt("Output clauses JSONL", "/data/staging/clauses-ecfr.jsonl")

    print("=== Step 2: Parsing ===")
    total = 0
    with open(output_clauses_path, "w", encoding="utf-8") as output:
        for filename in sorted(os.listdir(ecfr_json_directory)):
            if not filename.endswith(".json"):
                continue
            part = filename.split("part-")[-1].replace(".json", "")
            with open(os.path.join(ecfr_json_directory, filename), encoding="utf-8") as handle:
                data = json.load(handle)
            rows = []
            walk_sections(data, part, rows)
            for row in rows:
                output.write(json.dumps(row) + "\n")
            print(f"  {filename}: {len(rows)} clauses")
            total += len(rows)
    print(f"Done. {total} eCFR clauses -> {output_clauses_path}")
    print("Spot-check one clause against ecfr.gov before loading (P4-03 acceptance).")


if __name__ == "__main__":
    main()

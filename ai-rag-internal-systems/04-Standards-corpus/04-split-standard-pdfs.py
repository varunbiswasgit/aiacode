#!/usr/bin/env python3
"""Extract text from EU GMP / ICH PDFs and split it into clause records.

Tracker task: P4-06 (PDF text extraction + clause splitting)
Prompts: pdf_directories (comma-separated), output_clauses_path,
         split_pattern (regex for clause headings)
Prerequisites: Python 3.12. Optional: pypdf (used if installed; otherwise
  the script falls back to pdftotext if present on the system).
Output JSONL rows: standard, chapter, clause_no, text, version, source_url.
  version = the run date; diffing happens in 04-detect-standard-changes.py.
"""
import datetime
import json
import os
import re
import subprocess


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def extract_pdf_text(path):
    try:
        from pypdf import PdfReader
        return "\n".join((page.extract_text() or "") for page in PdfReader(path))
    except ImportError:
        pass
    result = subprocess.run(["pdftotext", "-layout", path, "-"],
                            capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout
    print(f"  WARNING: no PDF extractor available for {path} (pip install pypdf)")
    return ""


def main():
    print("=== Step 1: Split settings ===")
    pdf_directories = prompt("PDF directories, comma-separated (example — replace with your value)",
                             "/data/staging/eugmp,/data/staging/ich")
    output_clauses_path = prompt("Output clauses JSONL (example — replace with your value)", "/data/staging/clauses-standards.jsonl")
    split_pattern = prompt("Clause-heading regex (named group 'clause') (example — replace with your value)",
                           r"^(?P<clause>\d+(?:\.\d+)+)\s")
    try:
        heading = re.compile(split_pattern, re.MULTILINE)
    except re.error as error:
        print(f"ERROR: invalid regex: {error}")
        raise SystemExit(1)
    version = datetime.date.today().isoformat()

    print("=== Step 2: Extracting and splitting ===")
    total = 0
    with open(output_clauses_path, "w", encoding="utf-8") as output:
        for directory in [entry.strip() for entry in pdf_directories.split(",") if entry.strip()]:
            if not os.path.isdir(directory):
                print(f"  Skipping missing directory {directory}")
                continue
            for filename in sorted(os.listdir(directory)):
                if not filename.lower().endswith(".pdf"):
                    continue
                path = os.path.join(directory, filename)
                text = extract_pdf_text(path)
                if not text.strip():
                    continue
                standard = "EU GMP" if "eugmp" in filename else "ICH"
                matches = list(heading.finditer(text))
                if not matches:
                    output.write(json.dumps({
                        "standard": standard, "chapter": filename,
                        "clause_no": "full-document", "text": text.strip()[:8000],
                        "version": version, "source_url": "",
                    }) + "\n")
                    total += 1
                    continue
                for index, match in enumerate(matches):
                    end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
                    output.write(json.dumps({
                        "standard": standard, "chapter": filename,
                        "clause_no": match.group("clause"),
                        "text": text[match.start():end].strip(),
                        "version": version, "source_url": "",
                    }) + "\n")
                    total += 1
                print(f"  {filename}: clauses extracted")
    print(f"Done. {total} clause records -> {output_clauses_path}")


if __name__ == "__main__":
    main()

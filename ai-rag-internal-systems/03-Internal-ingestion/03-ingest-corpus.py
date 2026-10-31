#!/usr/bin/env python3
"""Run a full or incremental ingest: parse -> chunk -> enrich -> JSONL out.

Tracker tasks: P3-09 (first load / ingest pilot corpus),
               P10-04 (expand corpora, phase 2 sources - rerun with a new source)
Prompts: source_directory, corpus_name, chunk_size_words, chunk_overlap_words,
         output_jsonl_path, ingest_mode (full/incremental)
Prerequisites: Python 3.12. Reuses 03-parse-documents.py + 03-extract-metadata.py
  when they sit next to this script; otherwise falls back to plain-text parsing.
For vector-store loading, pipe the output JSONL into 05-embed-documents.py.
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from parse_documents import extract_text, chunk_words  # noqa: F401
    HAVE_PARSER = True
except ImportError:
    HAVE_PARSER = False


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
    print("=== Step 1: Ingest settings ===")
    source_directory = prompt("Source directory", "/data/staging/pilot")
    corpus_name = prompt("Corpus name (pilot, phase2, ...)", "pilot")
    chunk_size = prompt_int("Chunk size (words)", 750)
    chunk_overlap = prompt_int("Chunk overlap (words)", 150)
    output_jsonl_path = prompt("Output JSONL path", f"/data/staging/chunks-{corpus_name}.jsonl")
    ingest_mode = prompt("Mode (full/incremental)", "full").lower()
    if ingest_mode not in ("full", "incremental"):
        print("ERROR: mode must be full or incremental.")
        raise SystemExit(1)
    if not HAVE_PARSER:
        print("NOTE: 03-parse-documents.py not found next to this script;")
        print("      only .txt/.md files will be parsed.")

    print(f"=== Step 2: Ingesting corpus '{corpus_name}' ({ingest_mode}) ===")
    seen_paths = set()
    if ingest_mode == "incremental" and os.path.exists(output_jsonl_path):
        with open(output_jsonl_path, encoding="utf-8") as existing:
            for line in existing:
                seen_paths.add(json.loads(line).get("source_path", ""))
        print(f"Incremental mode: {len(seen_paths)} chunks already present, skipping unchanged files.")

    total_docs = 0
    total_chunks = 0
    mode = "a" if ingest_mode == "incremental" else "w"
    with open(output_jsonl_path, mode, encoding="utf-8") as output:
        for root, _dirs, files in os.walk(source_directory):
            for filename in sorted(files):
                path = os.path.join(root, filename)
                if ingest_mode == "incremental" and path in seen_paths:
                    continue
                if HAVE_PARSER:
                    text = extract_text(path)
                elif filename.lower().endswith((".txt", ".md")):
                    with open(path, encoding="utf-8", errors="ignore") as handle:
                        text = handle.read()
                else:
                    continue
                if not text.strip():
                    continue
                words = re.split(r"\s+", text.strip())
                doc_id = os.path.splitext(os.path.relpath(path, source_directory))[0]
                step = max(chunk_size - chunk_overlap, 1)
                for index in range(0, len(words), step):
                    chunk = " ".join(words[index:index + chunk_size])
                    output.write(json.dumps({
                        "corpus": corpus_name, "doc_id": doc_id,
                        "source_path": path, "chunk_index": index // step,
                        "text": chunk,
                    }) + "\n")
                    total_chunks += 1
                    if index + chunk_size >= len(words):
                        break
                total_docs += 1
    print(f"Done. {total_docs} documents, {total_chunks} new chunks -> {output_jsonl_path}")
    print("Next: 03-extract-metadata.py, then 05-embed-documents.py.")


if __name__ == "__main__":
    main()

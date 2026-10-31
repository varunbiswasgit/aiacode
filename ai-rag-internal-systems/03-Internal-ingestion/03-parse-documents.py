#!/usr/bin/env python3
"""Parse documents (PDF, DOCX, TXT, HTML) into chunked JSONL for the RAG pipeline.

Tracker task: P3-05 (document parsers)
Prompts: input_directory, output_jsonl_path, chunk_size_tokens (approx words),
         chunk_overlap_words
Prerequisites: Python 3.12. Optional: pypdf, python-docx (used if installed,
  otherwise PDF/DOCX files are skipped with a warning).
Output: one JSON object per line: doc_id, source_path, chunk_index, text.
"""
import html.parser
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET


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


def extract_pdf_text(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        print(f"  WARNING: pypdf not installed, skipping {path} (pip install pypdf)")
        return ""
    return "\n".join((page.extract_text() or "") for page in PdfReader(path))


def extract_docx_text(path):
    try:
        from docx import Document
        return "\n".join(paragraph.text for paragraph in Document(path).paragraphs)
    except ImportError:
        pass
    # Fallback: read word/document.xml straight from the zip.
    try:
        with zipfile.ZipFile(path) as archive:
            xml_bytes = archive.read("word/document.xml")
        root = ET.fromstring(xml_bytes)
        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        return "\n".join("".join(node.itertext())
                         for node in root.iterfind(".//w:p", namespace))
    except Exception as error:
        print(f"  WARNING: could not parse {path}: {error}")
        return ""


class HtmlTextExtractor(html.parser.HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        text = data.strip()
        if text:
            self.parts.append(text)


def extract_text(path):
    lower = path.lower()
    if lower.endswith(".pdf"):
        return extract_pdf_text(path)
    if lower.endswith(".docx"):
        return extract_docx_text(path)
    if lower.endswith((".txt", ".md")):
        with open(path, encoding="utf-8", errors="ignore") as handle:
            return handle.read()
    if lower.endswith((".html", ".htm")):
        extractor = HtmlTextExtractor()
        with open(path, encoding="utf-8", errors="ignore") as handle:
            extractor.feed(handle.read())
        return "\n".join(extractor.parts)
    if lower.endswith(".xlsx"):
        print(f"  NOTE: spreadsheets need manual handling, skipping {path}")
        return ""
    print(f"  NOTE: unsupported type, skipping {path}")
    return ""


def chunk_words(words, chunk_size, overlap):
    step = max(chunk_size - overlap, 1)
    for start in range(0, len(words), step):
        yield " ".join(words[start:start + chunk_size])
        if start + chunk_size >= len(words):
            break


def main():
    print("=== Step 1: Parse settings ===")
    input_directory = prompt("Input directory with documents", "/data/staging/pilot")
    output_jsonl_path = prompt("Output JSONL path", "/data/staging/chunks.jsonl")
    chunk_size = prompt_int("Chunk size (words, approx 1000 tokens ~ 750 words)", 750)
    chunk_overlap = prompt_int("Chunk overlap (words)", 150)

    print("=== Step 2: Parsing and chunking ===")
    total_docs = 0
    total_chunks = 0
    with open(output_jsonl_path, "w", encoding="utf-8") as output:
        for root, _dirs, files in os.walk(input_directory):
            for filename in sorted(files):
                path = os.path.join(root, filename)
                text = extract_text(path)
                if not text.strip():
                    continue
                words = re.split(r"\s+", text.strip())
                doc_id = os.path.splitext(os.path.relpath(path, input_directory))[0]
                for index, chunk in enumerate(chunk_words(words, chunk_size, chunk_overlap)):
                    output.write(json.dumps({
                        "doc_id": doc_id, "source_path": path,
                        "chunk_index": index, "text": chunk,
                    }) + "\n")
                    total_chunks += 1
                total_docs += 1
    print(f"Done. {total_docs} documents -> {total_chunks} chunks in {output_jsonl_path}")


if __name__ == "__main__":
    main()

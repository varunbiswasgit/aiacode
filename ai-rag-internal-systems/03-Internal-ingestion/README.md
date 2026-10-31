# 03 — Internal ingestion

## Purpose

Turn internal documents (SharePoint, SMB shares, Confluence) into
chunked, metadata-enriched JSONL ready for embedding. This folder covers
the scriptable ingestion steps; source inventory (P3-01), chunking-rule
design (P3-04) and the human mapping review (P3-12) stay manual.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 03-connect-document-sources.py | P3-02, P3-03 | Tests read access to one source (SharePoint Graph, SMB via smbclient, Confluence REST) and saves a sample listing. |
| 03-parse-documents.py | P3-05 | Parses PDF (pypdf), DOCX, TXT/Markdown and HTML into chunked JSONL. |
| 03-extract-metadata.py | P3-06 | Adds title, owner, department and a filename-derived doc ID to every chunk. |
| 03-install-ocr.sh | P3-07 | Installs Tesseract and ocrmypdf for scanned PDFs. |
| 03-schedule-ingestion.sh | P3-08 | Creates systemd service and timers for full (weekly) and incremental (hourly) ingestion. |
| 03-ingest-corpus.py | P3-09, P10-04 | Full or incremental corpus ingest: parse, chunk, write JSONL. Re-run with a new source directory for phase-2 expansion. |
| 03-validate-ingest.py | P3-10 | Shows random chunks and asks a human to confirm each; reports a pass rate and the docs needing fixes. |
| 03-map-documents-to-clauses.py | P3-11 | Maps chunks to clauses via explicit citations (21 CFR, ICH Q, EU GMP Annex/Chapter) plus keyword overlap; keyword-only matches are marked needs_review. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 03-connect-document-sources.py | source_type, source_url, username, access_token (secret), sample_output_path |
| 03-parse-documents.py | input_directory, output_jsonl_path, chunk_size, chunk_overlap |
| 03-extract-metadata.py | input_jsonl_path, output_jsonl_path, doc_id_pattern, default_owner, default_department |
| 03-install-ocr.sh | venv_path, ocr_languages |
| 03-schedule-ingestion.sh | ingest_command, run_user, project_dir, log_dir, full_schedule, incremental_schedule |
| 03-ingest-corpus.py | source_directory, corpus_name, chunk_size, chunk_overlap, output_jsonl_path, ingest_mode (full/incremental) |
| 03-validate-ingest.py | input_jsonl_path, sample_size, then y/n per sampled chunk |
| 03-map-documents-to-clauses.py | chunks_jsonl_path, clauses_jsonl_path, output_mappings_path, keyword_overlap_threshold |

Access tokens are read with getpass, never printed or saved. Only the
standard library is required for the Python scripts, except pypdf and
python-docx which are used opportunistically for PDF/DOCX parsing.

## Logic flow

Connect/test sources → parse → enrich metadata → (optional OCR) →
schedule → ingest → validate sample → map to clauses. The JSONL chunk
format is `{corpus, doc_id, source_path, chunk_index, text}` plus a
`metadata` object after enrichment; the same format feeds phase 05
embedding.

## Configuration

- Default chunk size: 750 words (about 1000 tokens), overlap 150 words.
- Default staging root: /data/staging.
- Default schedules: full re-index Sundays 02:00, incremental hourly.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

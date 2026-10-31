AI-RAG Internal Systems is an on-premises retrieval-augmented generation platform for internal content discovery, question answering over internal documents and regulatory standards (FDA eCFR Title 21, EU GMP, ICH), and document-update recommendations when a standard changes. It is generic: every deployment value is supplied as user input at runtime, and no company-specific value is hardcoded. Every recommendation is a suggestion for human QA and document-owner approval; the system never auto-edits a controlled document.

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
| 03-schedule-ingestion.sh | P3-08 | Creates systemd service and timers for full (weekly) and incremental (hourly) ingestion, with hardening flags (ProtectSystem=strict, UMask 0027, NoNewPrivileges, PrivateTmp) and a new extra_write_dirs prompt (default example /data/staging) listing every additional writable directory in ReadWritePaths. |
| 03-ingest-corpus.py | P3-09, P10-04 | Full or incremental corpus ingest: parse, chunk, write JSONL. Re-run with a new source directory for phase-2 expansion. |
| 03-validate-ingest.py | P3-10 | Shows random chunks and asks a human to confirm each; reports a pass rate and the docs needing fixes. |
| 03-map-documents-to-clauses.py | P3-11 | Maps chunks to clauses via explicit citations (21 CFR, ICH Q, EU GMP Annex/Chapter) plus keyword overlap; keyword-only matches are marked needs_review. |
| 03-build-chunks-fts.py | P3-13 | Builds the chunks_fts Postgres table from chunk JSONL for hybrid retrieval: upsert (safe to re-run, matched on doc_id + chunk_index) or rebuild (drops and recreates chunks_fts, only after an explicit confirmation). Rows missing doc_id, chunk_index or text are skipped and counted, never loaded. The keyword column is computed on the server with to_tsvector('english', text) and indexed with a GIN index (chunks_fts_tsv_idx). The password comes from PGPASSWORD or getpass only, never from argv, and is passed to psql only through the subprocess environment. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 03-connect-document-sources.py | source_type, source_url, username, access_token (secret), sample_output_path (default /opt/airag/source-sample.json, written mode 600) |
| 03-parse-documents.py | input_directory, output_jsonl_path, chunk_size, chunk_overlap |
| 03-extract-metadata.py | input_jsonl_path, output_jsonl_path, doc_id_pattern, default_owner, default_department |
| 03-install-ocr.sh | venv_path, ocr_languages |
| 03-schedule-ingestion.sh | ingest_command, run_user, project_dir, log_dir, extra_write_dirs (default example /data/staging), full_schedule, incremental_schedule |
| 03-ingest-corpus.py | source_directory, corpus_name, chunk_size, chunk_overlap, output_jsonl_path, ingest_mode (full/incremental) |
| 03-validate-ingest.py | input_jsonl_path, sample_size, then y/n per sampled chunk |
| 03-map-documents-to-clauses.py | chunks_jsonl_path, clauses_jsonl_path, output_mappings_path, keyword_overlap_threshold |
| 03-build-chunks-fts.py | chunks_jsonl_path, db_host, db_port, db_name, db_user, db_password (secret, PGPASSWORD or getpass, never argv), load_mode (upsert/rebuild) |

Access tokens are read with getpass, never printed or saved. SMB access
uses smbclient -A with an auth file created mode 600 and deleted
afterwards; the credential is never passed on the command line. Only the
standard library is required for the Python scripts, except pypdf and
python-docx which are used opportunistically for PDF/DOCX parsing.
Parse, ingest, metadata and mapping outputs are chmod 600. Chunking
prompts are chunk_size and chunk_overlap (words), not
chunk_size_words or chunk_size_tokens.

## Logic flow

Connect/test sources → parse → enrich metadata → (optional OCR) →
schedule → ingest → validate sample → map to clauses → build
chunks_fts (03-build-chunks-fts.py) before phase-05 retrieval. The JSONL
chunk format is `{corpus, doc_id, source_path, chunk_index, text}` plus a
`metadata` object after enrichment; the same format feeds phase 05
embedding. chunks_fts schema: doc_id TEXT NOT NULL, chunk_index INTEGER
NOT NULL, corpus, source_path, text TEXT NOT NULL, tsv tsvector NOT
NULL, PRIMARY KEY (doc_id, chunk_index), plus GIN index
chunks_fts_tsv_idx on tsv.

## Ordering notes

- The FTS build must run before 05-hybrid-retrieve.py: retrieval reads
  chunks_fts for its keyword path.
- Table-creating scripts run as the app user (this FTS build,
  07-populate-access-map.py) need an owner/privileged first run or
  pre-created tables now that schema CREATE is dropped from the app
  user's grants: the app user holds USAGE on the schema only.

## Configuration

- Default chunk size: 750 words (about 1000 tokens), overlap 150 words.
- Default staging root: /data/staging.
- Default schedules: full re-index Sundays 02:00, incremental hourly.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |
| v2 | Fix pass: new FTS build script P3-13 (upsert/rebuild from chunks JSONL, english tsvector + GIN index, PGPASSWORD only); connect: SMB via smbclient -A mode-600 auth file deleted afterwards (never argv), sample default /opt/airag/source-sample.json mode 600; ingest/parse/extract/map headers use chunk_size/chunk_overlap and outputs are chmod 600; schedule ingestion hardening flags + new extra_write_dirs prompt (default example /data/staging); other scripts: labels only. |

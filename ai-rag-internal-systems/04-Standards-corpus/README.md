AI-RAG Internal Systems is an on-premises retrieval-augmented generation platform for internal content discovery, question answering over internal documents and regulatory standards (FDA eCFR Title 21, EU GMP, ICH), and document-update recommendations when a standard changes. It is generic: every deployment value is supplied as user input at runtime, and no company-specific value is hardcoded. Every recommendation is a suggestion for human QA and document-owner approval; the system never auto-edits a controlled document.

# 04 — Standards corpus

## Purpose

Build and maintain the versioned clause registry for the three standards
families: FDA eCFR Title 21, EU GMP (EudraLex Volume 4) and ICH
guidelines (Q7, Q8, Q9, Q10, Q12). The registry is append-only: every
fetch creates a new dated version, and the change detector diffs fresh
fetches against the latest registered version, writing added, removed
and modified clauses to an audit table. The version-retention policy
(P4-11) is enforced by the schema itself (the app user gets no
UPDATE/DELETE on clause tables) and stays a QA-owned rule.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 04-fetch-ecfr.py | P4-01, P4-02 | Checks the eCFR API, then fetches the chosen Title 21 parts as JSON. |
| 04-parse-ecfr-clauses.py | P4-03 | Splits eCFR part JSON into individual clause records (clause JSONL). |
| 04-fetch-eugmp.py | P4-04 | Downloads EU GMP chapter/annex PDFs (URLs prompted, since the EC site reorganises links) and writes a manifest. |
| 04-fetch-ich.py | P4-05 | Downloads ICH guideline PDFs (URLs prompted) and writes a manifest. |
| 04-split-standard-pdfs.py | P4-06 | Extracts PDF text (pypdf or pdftotext) and splits it into clause records by heading pattern. |
| 04-create-clause-registry.sql | P4-07 | Creates the clauses and clause_changes tables with the uniqueness rule on (standard, chapter, clause_no, version), and issues real INSERT/SELECT grants to the app user (no UPDATE/DELETE). Requires -v standards_db and -v app_user, and must run after 02-create-databases.sql creates the app user, as the table owner or a superuser so the GRANT statements succeed. |
| 04-load-standards-baseline.py | P4-08 | Loads clause JSONL into the registry via a temp staging table and upsert (safe to re-run). |
| 04-detect-standard-changes.py | P4-09 | Compares a fresh clause JSONL against the latest registered versions and records added/removed/modified changes. |
| 04-schedule-standards-diff.sh | P4-09 | Creates the systemd service and timer for scheduled change detection, with hardening flags (ProtectSystem=strict, UMask 0027, NoNewPrivileges, PrivateTmp) and a new extra_write_dirs prompt (default example /data/staging) listing every additional writable directory in ReadWritePaths. |
| 04-send-change-alert.py | P4-10 | Emails and/or posts a Teams alert listing today's clause changes and the ACTUAL affected documents, found by joining each change to the mappings file (no link to a nonexistent view is included). Impact ratings come from an optional suggestions JSONL; without one the alert prints "not assessed" instead of inventing a rating. Note: impact ratings from the suggestions database table surface in the app's /updates page instead. The temporary SQL file is deleted in try/finally, including on failure. |
| 04-generate-suggestions.py | P4-12 | Generates document-update suggestions from clause changes: for each pending clause_changes row it finds affected documents via mappings.jsonl, runs the gap.txt then suggest.txt templates against a local model, parses the impact rating (high/medium/low, never guessed or defaulted — a document with no recognisable rating or an empty suggestion is skipped with no row written), and INSERTs one suggestions row per (change_id, document_id, clause_id) with status pending. Idempotent: a change with any existing suggestion row is skipped whole. It CREATEs the suggestions table if missing and never edits a controlled document. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 04-fetch-ecfr.py | ecfr_title, ecfr_parts, output_directory, http_proxy (optional) |
| 04-parse-ecfr-clauses.py | ecfr_json_directory, output_clauses_path |
| 04-fetch-eugmp.py | eugmp_chapters, eugmp_annexes, output_directory, http_proxy (optional), pdf_url per chapter/annex (blank = skip that item) |
| 04-fetch-ich.py | ich_guidelines, output_directory, http_proxy (optional), pdf_url per guideline (blank = skip that item) |
| 04-split-standard-pdfs.py | pdf_directories, output_clauses_path, split_pattern |
| 04-create-clause-registry.sql | standards_db |
| 04-load-standards-baseline.py | clauses_jsonl_paths, db_host, db_port, db_name, db_user, db_password (secret) |
| 04-detect-standard-changes.py | new_clauses_jsonl_path, db_host, db_port, db_name, db_user, db_password (secret), alert_on_change |
| 04-schedule-standards-diff.sh | diff_command, run_user, project_dir, log_dir, extra_write_dirs (default example /data/staging), diff_schedule |
| 04-send-change-alert.py | db_host, db_port, db_name, db_user, db_password (secret), mappings_jsonl_path, suggestions_jsonl_path (optional, blank = no ratings file), alert_channel, smtp_relay_host, smtp_relay_port, alert_from_address, alert_to_addresses, teams_webhook_url (secret), send_test_only |
| 04-generate-suggestions.py | db_host, db_port, db_name, db_user, db_password (secret, PGPASSWORD or getpass, never argv), mappings_path, chunks_path (optional, blank = no section text supplied), templates_directory, model_endpoint, model_name, model_api_style (ollama/openai), model_timeout_seconds (default 120), detected_since (optional YYYY-MM-DD, blank = no window), batch_limit (default 20), include_needs_review_mappings (yes/no, default no) |

Database passwords and the Teams webhook URL are read with getpass,
never printed or saved. The eCFR API is free and needs no key.

## Logic flow

Fetch (eCFR API, EU GMP PDFs, ICH PDFs) → parse/split into clause
records → create registry → baseline load → scheduled re-fetch →
detect changes → generate suggestions → alert QA. Clause JSONL rows are
`{standard, chapter, clause_no, clause_path, text, version, source_url}`:
clause_no is the LEAF label only (for example 211.100), and clause_path
keeps the full ancestor path for traceability; clause_path is never a
join key. The baseline-load temporary SQL file is deleted in
try/finally, including on failure.

## Ordering notes

- The clause registry (04-create-clause-registry.sql) must run after
  02-create-databases.sql creates the app user (see phase 02).
- The suggestion step (04-generate-suggestions.py) runs after change
  detection (04-detect-standard-changes.py) and needs mappings.jsonl
  from 03-map-documents-to-clauses.py. It creates the suggestions table
  (suggestion_id, change_id, document_id, clause_id, impact_rating
  high/medium/low, suggested_text, status default pending, reviewer,
  comment, created_at) in the standards database; the grants script
  (02-grant-database-permissions.sh) must then be re-run so the app
  user gains INSERT/SELECT on suggestions.
- Alert impact ratings come from the optional suggestions JSONL, while
  impact ratings from the suggestions database table surface in the
  app's /updates page.

## Configuration

- Default eCFR title: 21; default parts: 210, 211, 11.
- Default EU GMP chapters: 1-9; annexes: 1, 11, 15.
- Default ICH guidelines: Q7, Q8, Q9, Q10, Q12.
- Default diff schedule: daily 02:00.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |
| v2 | Fix pass: new suggestion script P4-12 (clause_changes x mappings.jsonl, gap/suggest templates via local model, impact ratings, suggestions store, idempotent, never edits controlled docs); clause registry issues real INSERT/SELECT grants and must run after 02 creates the app user; eCFR clause_no is now the leaf label only with clause_path kept for traceability (never a join key); baseline-load temp file deleted on failure too (try/finally); scheduled diff hardening + extra_write_dirs prompt; alert lists actual affected documents via the mappings join with ratings from the optional suggestions JSONL ("not assessed" without one) and temp cleanup; fetch headers list per-item pdf_url prompts; other scripts: labels only. |

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
| 04-create-clause-registry.sql | P4-07 | Creates the clauses and clause_changes tables with the uniqueness rule on (standard, chapter, clause_no, version). |
| 04-load-standards-baseline.py | P4-08 | Loads clause JSONL into the registry via a temp staging table and upsert (safe to re-run). |
| 04-detect-standard-changes.py | P4-09 | Compares a fresh clause JSONL against the latest registered versions and records added/removed/modified changes. |
| 04-schedule-standards-diff.sh | P4-09 | Creates the systemd service and timer for scheduled change detection. |
| 04-send-change-alert.py | P4-10 | Emails and/or posts a Teams alert listing today's clause changes. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 04-fetch-ecfr.py | ecfr_title, ecfr_parts, output_directory, http_proxy (optional) |
| 04-parse-ecfr-clauses.py | ecfr_json_directory, output_clauses_path |
| 04-fetch-eugmp.py | eugmp_chapters, eugmp_annexes, output_directory, http_proxy (optional), per-item PDF URL |
| 04-fetch-ich.py | ich_guidelines, output_directory, http_proxy (optional), per-guideline PDF URL |
| 04-split-standard-pdfs.py | pdf_directories, output_clauses_path, split_pattern |
| 04-create-clause-registry.sql | standards_db |
| 04-load-standards-baseline.py | clauses_jsonl_paths, db_host, db_port, db_name, db_user, db_password (secret) |
| 04-detect-standard-changes.py | new_clauses_jsonl_path, db_host, db_port, db_name, db_user, db_password (secret), alert_on_change |
| 04-schedule-standards-diff.sh | diff_command, run_user, project_dir, log_dir, diff_schedule |
| 04-send-change-alert.py | db_host, db_port, db_name, db_user, db_password (secret), alert_channel, smtp_relay_host, smtp_relay_port, alert_from_address, alert_to_addresses, teams_webhook_url (secret), send_test_only |

Database passwords and the Teams webhook URL are read with getpass,
never printed or saved. The eCFR API is free and needs no key.

## Logic flow

Fetch (eCFR API, EU GMP PDFs, ICH PDFs) → parse/split into clause
records → create registry → baseline load → scheduled re-fetch →
detect changes → alert QA. Clause JSONL rows are
`{standard, chapter, clause_no, text, version, source_url}`.

## Configuration

- Default eCFR title: 21; default parts: 210, 211, 11.
- Default EU GMP chapters: 1-9; annexes: 1, 11, 15.
- Default ICH guidelines: Q7, Q8, Q9, Q10, Q12.
- Default diff schedule: daily 02:00.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

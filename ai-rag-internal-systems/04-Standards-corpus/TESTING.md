# Testing — 04 Standards corpus

Manual test cases. Run each case, then tick the Pass column. Use test
database names so the production registry is never touched by tests.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| App test node with network route to ecfr.gov | Run 04-fetch-ecfr.py for title 21 part 211 only | Script reports the API check OK and saves one part JSON file | ☐ |
| Part JSON from the fetch test | Run 04-parse-ecfr-clauses.py | Output JSONL has clause rows with standard, chapter, clause_no, clause_path, text, version; clause_no is the short leaf label only (for example 211.100, not the full path) and clause_path holds the full ancestor path; spot-check one clause against ecfr.gov text | ☐ |
| App test node with network route to the EC site | Run 04-fetch-eugmp.py for one chapter and one annex, pasting current PDF URLs | PDFs saved, manifest.json lists each downloaded file and URL | ☐ |
| App test node with network route to ich.org | Run 04-fetch-ich.py for Q7 only, pasting the current PDF URL | PDF saved, manifest.json lists it | ☐ |
| One EU GMP or ICH PDF from the fetch tests | Run 04-split-standard-pdfs.py on that directory | Output JSONL has clause rows split at headings, or one full-document row when no headings match | ☐ |
| Test Postgres database | Run 04-create-clause-registry.sql via psql with a test database name | clauses and clause_changes tables exist with the uniqueness constraint | ☐ |
| Test database with registry tables, small clause JSONL | Run 04-load-standards-baseline.py, then run it a second time | First load inserts rows, second load inserts zero duplicates, per-standard counts print correctly | ☐ |
| Test database with a baseline, plus a modified clause JSONL | Run 04-detect-standard-changes.py | clause_changes gains rows of the right change types, fresh clauses stored as new versions | ☐ |
| App test node, detect script present | Run 04-schedule-standards-diff.sh with a test schedule, answering the extra_write_dirs prompt with every directory the diff run writes to (default example /data/staging) | systemctl list-timers shows the airag-standards-diff timer enabled; the service uses ProtectSystem=strict with ReadWritePaths covering the project dir, log dir and the prompted extra_write_dirs, so a test output path not entered there is not writable by the scheduled run | ☐ |
| Test database with a change recorded today | Run 04-send-change-alert.py with send_test_only set to yes, answering the mappings-path prompt and leaving the suggestions-path prompt blank | Test alert arrives by the chosen channel listing the actual affected documents from the mappings file, with impact shown as "not assessed" when no suggestions file is given; no secrets appear in the output and the temporary SQL file is deleted | ☐ |
| Test standards database with a pending clause change, small mappings JSONL, gap.txt/suggest.txt templates and a test local model endpoint | Run 04-generate-suggestions.py | One suggestions row per affected document with status pending and a high/medium/low impact rating; a second run writes no new rows for the same change; no controlled document file is modified | ☐ |

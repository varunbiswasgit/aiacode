# Testing — 04 Standards corpus

Manual test cases. Run each case, then tick the Pass column. Use test
database names so the production registry is never touched by tests.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| App test node with network route to ecfr.gov | Run 04-fetch-ecfr.py for title 21 part 211 only | Script reports the API check OK and saves one part JSON file | ☐ |
| Part JSON from the fetch test | Run 04-parse-ecfr-clauses.py | Output JSONL has clause rows with standard, chapter, clause_no, text, version; spot-check one clause against ecfr.gov text | ☐ |
| App test node with network route to the EC site | Run 04-fetch-eugmp.py for one chapter and one annex, pasting current PDF URLs | PDFs saved, manifest.json lists each downloaded file and URL | ☐ |
| App test node with network route to ich.org | Run 04-fetch-ich.py for Q7 only, pasting the current PDF URL | PDF saved, manifest.json lists it | ☐ |
| One EU GMP or ICH PDF from the fetch tests | Run 04-split-standard-pdfs.py on that directory | Output JSONL has clause rows split at headings, or one full-document row when no headings match | ☐ |
| Test Postgres database | Run 04-create-clause-registry.sql via psql with a test database name | clauses and clause_changes tables exist with the uniqueness constraint | ☐ |
| Test database with registry tables, small clause JSONL | Run 04-load-standards-baseline.py, then run it a second time | First load inserts rows, second load inserts zero duplicates, per-standard counts print correctly | ☐ |
| Test database with a baseline, plus a modified clause JSONL | Run 04-detect-standard-changes.py | clause_changes gains rows of the right change types, fresh clauses stored as new versions | ☐ |
| App test node, detect script present | Run 04-schedule-standards-diff.sh with a test schedule | systemctl list-timers shows the airag-standards-diff timer enabled | ☐ |
| Test database with a change recorded today | Run 04-send-change-alert.py with send_test_only set to yes | Test alert arrives by the chosen channel, no secrets appear in the output | ☐ |

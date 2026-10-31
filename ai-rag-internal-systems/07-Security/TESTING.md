# Testing — 07 Security

Manual test cases. Run each case, then tick the Pass column. Use test
credentials and test databases throughout.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| App test node | Run 07-configure-sso.sh with test OIDC values and a dummy secret | Config file created with mode 600, secret value not visible in the log | ☐ |
| Test Postgres database with psycopg installed | Run 07-populate-access-map.py with two test groups and prefixes, answering the loop prompts (group, then comma-separated prefixes) | access-map.json written, access_map table rows match the JSON; no temp SQL file is created. NOTE: psycopg is now a prerequisite, not just psql | ☐ |
| Test Postgres database | Run 07-create-audit-log.sql via psql with a test database name | audit_log, feedback and review_decisions tables exist | ☐ |
| Chunk JSONL containing a test SSN and email | Run 07-scan-pii.py, accepting the quarantine default (/opt/airag/quarantine/chunks-quarantine.jsonl) | Flagged chunks land in the quarantine file with pii_flags, clean file has the rest, both outputs are mode 600 and the quarantine directory is mode 700 when created by the script. NOTE: the scan is pattern-based only, not a guarantee | ☐ |
| App test node, firewall rules applied | Run 07-test-egress.sh with test URLs and question | Allowed host resolves, blocked host unreachable, app Q and A succeeds | ☐ |
| App test node reachable from the scanner host | Run 07-run-vulnerability-scan.sh with nmap-basic against the test host | Report file written, findings listed for triage | ☐ |
| Restored DB copy on an isolated network | Run 07-restore-drill.sh with test hosts and production reference counts | Script prints restored counts and a pass or fail verdict against the references | ☐ |

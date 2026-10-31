AI-RAG Internal Systems is an on-premises retrieval-augmented generation platform for internal content discovery, question answering over internal documents and regulatory standards (FDA eCFR Title 21, EU GMP, ICH), and document-update recommendations when a standard changes. It is generic: every deployment value is supplied as user input at runtime, and no company-specific value is hardcoded. Every recommendation is a suggestion for human QA and document-owner approval; the system never auto-edits a controlled document.

# 07 — Security

## Purpose

Lock the platform down: SSO configuration, AD-group access mapping,
append-only audit logging, PII scanning at ingest, egress verification,
vulnerability scanning, and backup-restore verification. The QA
sign-off on the "suggest, don't edit" control (P7-08) stays a manual
sign-off step.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 07-configure-sso.sh | P7-01 | Writes the app SSO config (OIDC or LDAP) with credentials prompted and the file locked to mode 600. |
| 07-populate-access-map.py | P7-02 | Interactively builds the AD-group to doc-prefix map, saves it as JSON for the permission filter, and writes it to the access_map table. The WHOLE insert uses parameterized psycopg queries (%s placeholders; group and prefix list passed as parameters, never interpolated into SQL and never hand-built), the database password reaches libpq via the PGPASSWORD environment variable only, and NO temp SQL file is created at all. |
| 07-create-audit-log.sql | P7-03 | Creates audit_log, feedback and review_decisions tables (append-only by grant design) and issues real INSERT/SELECT grants to the app user (no UPDATE/DELETE). Requires -v audit_db and -v app_user, and must run after 02-create-databases.sql creates the app user, as the table owner or a superuser so the GRANT statements succeed. |
| 07-scan-pii.py | P7-04 | Pattern-scans chunks for SSNs, card-like sequences and emails; quarantines hits, passes the rest. Both outputs are chmod 600; the quarantine directory is created mode 700 when it does not exist. This is a pattern-based scan only, not a guarantee: patterns can miss PII (unusual formats, obfuscation) and can flag non-PII, so quarantined documents need human review before release. |
| 07-test-egress.sh | P7-05 | Verifies the allowed standards host resolves, an external host is blocked, and app Q&A still works. |
| 07-run-vulnerability-scan.sh | P7-06 | Runs a scanner (Nessus/Qualys/OpenVAS via the company scanner, or a basic nmap vuln sweep) and records the report path for triage. |
| 07-restore-drill.sh | P7-07 | Compares Postgres and Qdrant counts on a restored copy against production references and records a pass/fail verdict. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 07-configure-sso.sh | sso_type (oidc/ldap), provider_url, config_path, client_id, client_secret (secret, OIDC only), ldap_bind_dn, ldap_bind_password (secret, LDAP only), test_username |
| 07-populate-access-map.py | db_host, db_port, db_name, db_user, db_password (secret), access_map_path, then per-group prefixes in a loop |
| 07-create-audit-log.sql | audit_db |
| 07-scan-pii.py | input_jsonl_path, output_clean_path, output_quarantine_path (default /opt/airag/quarantine/chunks-quarantine.jsonl; the /opt/airag/quarantine directory is created mode 700 when missing) |
| 07-test-egress.sh | app_base_url, test_question, allowed_proxy_host, blocked_test_host |
| 07-run-vulnerability-scan.sh | scanner_type, target_host, target_port, report_path, scanner_host |
| 07-restore-drill.sh | restored_db_host, restored_db_port, restored_db_name, restored_db_user, restored_db_password (secret), restored_qdrant_url, prod_doc_count, prod_collection_count |

All secrets use `read -s` or getpass and are never echoed, logged or
written to files. The PII scan is pattern-based and not a guarantee;
quarantined documents need human review before release.

## Logic flow

SSO config → access map (QA reviews and signs off) → audit tables →
PII scan (clean file only goes to the vector store) → egress test →
vulnerability scan and triage → restore drill. Audit tables are
append-only: grant the app user INSERT and SELECT only, never UPDATE
or DELETE.

## Ordering notes

- The audit-log SQL (07-create-audit-log.sql) must run after
  02-create-databases.sql creates the app user (see phase 02).
- The grants script (02-grant-database-permissions.sh) must be re-run
  after the suggestion step creates the suggestions table; it also
  covers the audit tables (audit_log, feedback, review_decisions).
- Table-creating scripts run as the app user (07-populate-access-map.py,
  03-build-chunks-fts.py) need an owner/privileged first run or
  pre-created tables now that schema CREATE is dropped from the app
  user's grants: the app user holds USAGE on the schema only.
- The vulnerability-scan header is fixed: nmap-basic is annotated as
  the basic option and scanner_host is prompted conditionally (in the
  nmap-basic branch).

## Configuration

- Default audit database: airag_audit.
- Default access map path: /opt/airag/access-map.json.
- Default SSO config path: /opt/airag/sso-config.yaml (mode 600).

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |
| v2 | Fix pass: SSO labels/logs with config_path in header order; audit-log SQL issues real INSERT/SELECT grants and must run after 02 creates the app user; access map uses parameterized psycopg inserts for the whole insert, PGPASSWORD env only, no temp SQL file anymore (header lists loop prompts group/prefixes); PII scan outputs chmod 600, quarantine default /opt/airag/quarantine (directories created mode 700), pattern-based limitation documented; restore/vuln-scan/egress: labels/logs, vuln-scan header fixed (nmap-basic annotation, scanner_host conditional). |

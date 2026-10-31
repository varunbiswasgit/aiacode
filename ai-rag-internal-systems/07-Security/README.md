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
| 07-populate-access-map.py | P7-02 | Interactively builds the AD-group to doc-prefix map, saves it as JSON for the permission filter, and writes it to the access_map table. |
| 07-create-audit-log.sql | P7-03 | Creates audit_log, feedback and review_decisions tables (append-only by grant design). |
| 07-scan-pii.py | P7-04 | Pattern-scans chunks for SSNs, card-like sequences and emails; quarantines hits, passes the rest. |
| 07-test-egress.sh | P7-05 | Verifies the allowed standards host resolves, an external host is blocked, and app Q&A still works. |
| 07-run-vulnerability-scan.sh | P7-06 | Runs a scanner (Nessus/Qualys/OpenVAS via the company scanner, or a basic nmap vuln sweep) and records the report path for triage. |
| 07-restore-drill.sh | P7-07 | Compares Postgres and Qdrant counts on a restored copy against production references and records a pass/fail verdict. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 07-configure-sso.sh | sso_type (oidc/ldap), provider_url, client_id, client_secret (secret, OIDC), ldap_bind_dn, ldap_bind_password (secret, LDAP), config_path, test_username |
| 07-populate-access-map.py | db_host, db_port, db_name, db_user, db_password (secret), access_map_path, then per-group prefixes in a loop |
| 07-create-audit-log.sql | audit_db |
| 07-scan-pii.py | input_jsonl_path, output_clean_path, output_quarantine_path |
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

## Configuration

- Default audit database: airag_audit.
- Default access map path: /opt/airag/access-map.json.
- Default SSO config path: /opt/airag/sso-config.yaml (mode 600).

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

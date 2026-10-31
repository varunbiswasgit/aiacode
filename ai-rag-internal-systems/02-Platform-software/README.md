AI-RAG Internal Systems is an on-premises retrieval-augmented generation platform for internal content discovery, question answering over internal documents and regulatory standards (FDA eCFR Title 21, EU GMP, ICH), and document-update recommendations when a standard changes. It is generic: every deployment value is supplied as user input at runtime, and no company-specific value is hardcoded. Every recommendation is a suggestion for human QA and document-owner approval; the system never auto-edits a controlled document.

# 02 — Platform software

## Purpose

Install the runtime platform on the three nodes: Docker, Python, Git,
PostgreSQL 16 with pgvector, Qdrant, Redis, Ollama, vLLM, Nginx with
internal-CA TLS, log aggregation, and a final smoke test. After this
phase the platform is ready for the ingestion and standards pipelines.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 02-install-docker.sh | P2-01, P2-12 | Installs Docker Engine and Compose (distro package or installer script, prompted), optionally configures an internal registry mirror. No service account is added to the docker group; verification uses sudo docker. |
| 02-install-python.sh | P2-02 | Installs Python 3.12 and creates the project venv. |
| 02-clone-project-repo.sh | P2-03 | Installs Git and clones (or pulls) the project repo onto the node. |
| 02-install-postgres.sh | P2-04 | Installs PostgreSQL 16 and enables the service. |
| 02-enable-pgvector.sh | P2-05 | Installs postgresql-16-pgvector and creates the extension in a database. |
| 02-create-databases.sql | P2-06 | Creates airag_docs, airag_standards, airag_audit and the scoped app user. Schema grants are USAGE only (CREATE is dropped), so the app user cannot create tables; tables are created by the phase scripts running as the table owner or a superuser. Run pattern: read the application-user password silently into PGPASSWORD with read -s, then run psql with -v docs_db, standards_db, audit_db and app_user. Passing the password via psql -v app_password is FORBIDDEN (it would expose it in the process list). |
| 02-install-qdrant.sh | P2-07 | Runs Qdrant as a container with a persistent volume and verifies the health endpoint. The API key is read with read -s, written to a mode-600 env file and passed with --env-file; ports are published only on the prompted internal IP (0.0.0.0 is rejected); the image must be digest-pinned (untagged or latest references are rejected); an existing container is recreated, with the volume preserved. |
| 02-install-redis.sh | P2-08 | Runs Redis as a container with append-only persistence and verifies PONG. The password is supplied with --requirepass, read with read -s and written to a mode-600 config file; verification uses REDISCLI_AUTH, not a command-line argument. The port is published only on the prompted internal IP (0.0.0.0 is rejected); the image must be digest-pinned (mutable tags such as redis:7 and latest are rejected); an existing container is recreated, with the volume preserved. |
| 02-install-ollama.sh | P2-09 | Installs Ollama on the GPU node and verifies the API. The installer is downloaded to a local file, SHA256-checked against the operator-supplied expected hash, and run only after a yes confirmation; it is never piped to a shell. The bind IP is prompted and 0.0.0.0 is rejected. |
| 02-install-vllm.sh | P2-10 | Installs vLLM into the project venv and verifies the import. The version is pinned (exact version only, never blank, latest or a range; default example 0.31.0); an optional hash-verified lockfile may be installed with pip --require-hashes instead, and a missing lockfile or a lockfile whose vLLM pin differs aborts the install. After install the script asserts installed version == pinned version and writes a pip freeze lockfile. |
| 02-install-nginx-tls.sh | P2-11 | Installs Nginx, creates a key and CSR, pauses for internal-CA signing, then writes the HTTPS reverse-proxy site. TLS is 1.2 and 1.3 only with modern ECDHE ciphers and an HSTS header; verification runs WITHOUT -k as the primary check (certificate validated against the system trust store), with the -k check retained only as a marked pre-CA fallback; the CSR is deleted once the signed certificate exists. |
| 02-configure-log-aggregation.sh | P2-13 | Forwards system logs and Docker logs to the SIEM via rsyslog/syslog driver. |
| 02-platform-smoke-test.sh | P2-14 | Checks Docker, Qdrant, Redis and Postgres reachability from the app node. |
| 02-grant-database-permissions.sh | P2-15 | Grants exactly INSERT and SELECT on the AI-RAG tables to the app user and any additional prompted usernames (standards: clauses, clause_changes; audit: audit_log, feedback, review_decisions; suggestions database: suggestions). No UPDATE or DELETE is granted. A table that does not exist yet is warned about and skipped, so the script must be re-run after the suggestion step creates the suggestions table. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 02-install-docker.sh | install_method (package/installer), installer_sha256_checksum (installer method only), use_internal_registry (yes/no), internal_registry_url |
| 02-install-python.sh | python_version, venv_path, venv_base, project_owner_account |
| 02-clone-project-repo.sh | git_server_hostname, git_repository_path, project_dir, project_owner_account, checkout_branch |
| 02-install-postgres.sh | postgres_version, postgres_data_dir |
| 02-enable-pgvector.sh | postgres_version, target_database |
| 02-create-databases.sql | docs_db, standards_db, audit_db, app_user (via -v); app password via PGPASSWORD only (read -s by the shell; never via -v app_password) |
| 02-install-qdrant.sh | qdrant_bind_ip, qdrant_http_port, qdrant_grpc_port, qdrant_volume_name, qdrant_image (digest-pinned), qdrant_api_key (secret, read -s) |
| 02-install-redis.sh | redis_bind_ip, redis_port, redis_volume_name, redis_image (digest-pinned), redis_password (secret, read -s) |
| 02-install-ollama.sh | ollama_bind_ip, ollama_port, ollama_models_dir, ollama_installer_sha256 (only when Ollama is not already installed) |
| 02-install-vllm.sh | venv_path, vllm_version (pinned, exact version only), vllm_lockfile_path (optional hash-verified lockfile) |
| 02-install-nginx-tls.sh | app_domain_name, app_upstream_port, cert_country, cert_org |
| 02-configure-log-aggregation.sh | siem_hostname_or_ip, siem_syslog_port, docker_log_driver |
| 02-platform-smoke-test.sh | db_node_hostname, qdrant_port, redis_port, postgres_port, qdrant_api_key (optional, read -s, blank = try without) |
| 02-grant-database-permissions.sh | db_host, db_port, db_admin_user, db_admin_password (secret, read -s into PGPASSWORD only, unset on exit), standards_db, audit_db, suggestions_db, app_user, additional_grantees (optional, space-separated) |

The database password is never written to disk by these scripts: for
02-create-databases.sql it is read silently into PGPASSWORD with
read -s and passed to psql only through that environment variable,
never via -v app_password. Pick ONE production model server (Ollama or
vLLM), not both.

## Logic flow

Docker → Python venv → Git clone → Postgres → pgvector → databases →
Qdrant → Redis → Ollama and/or vLLM → Nginx/TLS → log aggregation →
smoke test, then the grants script (P2-15). Container scripts are
idempotent (Qdrant and Redis existing containers are recreated, with
their volumes preserved, so a new image, key or bind takes effect).
Logs go to `/var/log/airag/` (mode 640, umask 027).

## Ordering notes

- 02-create-databases.sql must run before the table-creating scripts
  that grant to the app user: 04-create-clause-registry.sql and
  07-create-audit-log.sql both require -v app_user and must run after
  02 creates that user, and both issue real INSERT/SELECT grants (no
  longer echo hints). Run them as the table owner or a superuser so
  the GRANT statements succeed.
- The grants script (02-grant-database-permissions.sh) must be re-run
  after the suggestion step (04-generate-suggestions.py) creates the
  suggestions table; until then that table is warned about and skipped.
- Table-creating scripts run as the app user (03-build-chunks-fts.py,
  07-populate-access-map.py) need an owner/privileged first run or
  pre-created tables now that schema CREATE is dropped from the app
  user's grants: the app user holds USAGE on the schema only.
- 02-enable-pgvector.sh prints === step headers. Note, corrected: the
  docs and standards databases get the vector extension from
  02-create-databases.sql; the audit database does NOT get vector
  from either script. Whether the audit database needs vector is an
  open question — do not claim otherwise.
- The smoke test's Qdrant check takes an optional API key prompt:
  blank means try without a key, and a 401 answer prints a distinct
  auth-required message. That still counts as a failed check; it is
  not a reachability pass.
- Python venv creation takes a new venv_base prompt and refuses any
  venv whose parent is outside that base (and refuses a `/` base and
  `..` components), so a mistyped venv path cannot hand a system
  directory to the service account.
- Docker install takes a new install_method prompt (package/installer):
  with the installer method the published SHA256 checksum is verified
  before the downloaded local copy is run, and the installer is never
  piped to a shell.

## Configuration

- Default Postgres version: 16. Default Qdrant ports: 6333/6334.
  Default Redis port: 6379. Default Ollama port: 11434.
- Default databases: airag_docs, airag_standards, airag_audit.
- Default venv path: /opt/airag/venv. Default project dir: /opt/airag.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |
| v2 | Fix pass: Qdrant/Redis API key or password (read -s, mode-600 file), internal-only bind, digest-pinned image, existing containers recreated; Docker install_method prompt with installer checksum verified before running the local copy and docker-group prompts removed; Nginx TLSv1.2+ only, modern ciphers, HSTS, primary verification without -k with marked pre-CA -k fallback, CSR deleted after the signed cert exists; Python venv_base prefix validation; vLLM pinned version (default example 0.31.0), optional hash lockfile, installed==pinned assertion, freeze lockfile written; Ollama installer SHA256-checked with yes confirmation and prompted bind IP; databases SQL schema grants USAGE only with the PGPASSWORD run pattern (psql -v app_password forbidden); new grants script P2-15; smoke test optional Qdrant API key prompt with distinct 401 message; pgvector step headers and corrected vector note (audit DB has no vector). |

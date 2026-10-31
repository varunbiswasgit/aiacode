# 02 — Platform software

## Purpose

Install the runtime platform on the three nodes: Docker, Python, Git,
PostgreSQL 16 with pgvector, Qdrant, Redis, Ollama, vLLM, Nginx with
internal-CA TLS, log aggregation, and a final smoke test. After this
phase the platform is ready for the ingestion and standards pipelines.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 02-install-docker.sh | P2-01, P2-12 | Installs Docker Engine and Compose via get.docker.com, adds the service account to the docker group, optionally configures an internal registry mirror. |
| 02-install-python.sh | P2-02 | Installs Python 3.12 and creates the project venv. |
| 02-clone-project-repo.sh | P2-03 | Installs Git and clones (or pulls) the project repo onto the node. |
| 02-install-postgres.sh | P2-04 | Installs PostgreSQL 16 and enables the service. |
| 02-enable-pgvector.sh | P2-05 | Installs postgresql-16-pgvector and creates the extension in a database. |
| 02-create-databases.sql | P2-06 | Creates airag_docs, airag_standards, airag_audit and the scoped app user (password prompted, never stored in the file). |
| 02-install-qdrant.sh | P2-07 | Runs Qdrant as a container with a persistent volume and verifies the health endpoint. |
| 02-install-redis.sh | P2-08 | Runs Redis as a container with append-only persistence and verifies PONG. |
| 02-install-ollama.sh | P2-09 | Installs Ollama on the GPU node and verifies the API. |
| 02-install-vllm.sh | P2-10 | Installs vLLM into the project venv and verifies the import. |
| 02-install-nginx-tls.sh | P2-11 | Installs Nginx, creates a key and CSR, pauses for internal-CA signing, then writes the HTTPS reverse-proxy site. |
| 02-configure-log-aggregation.sh | P2-13 | Forwards system logs and Docker logs to the SIEM via rsyslog/syslog driver. |
| 02-platform-smoke-test.sh | P2-14 | Checks Docker, Qdrant, Redis and Postgres reachability from the app node. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 02-install-docker.sh | docker_user_account, use_internal_registry (yes/no), internal_registry_url |
| 02-install-python.sh | python_version, venv_path, project_owner_account |
| 02-clone-project-repo.sh | git_server_hostname, git_repository_path, project_dir, project_owner_account, checkout_branch |
| 02-install-postgres.sh | postgres_version, postgres_data_dir |
| 02-enable-pgvector.sh | postgres_version, target_database |
| 02-create-databases.sql | docs_db, standards_db, audit_db, app_user, app_password (secret, read silently by the shell wrapper and passed via -v) |
| 02-install-qdrant.sh | qdrant_http_port, qdrant_grpc_port, qdrant_volume_name, qdrant_image |
| 02-install-redis.sh | redis_port, redis_volume_name, redis_image |
| 02-install-ollama.sh | ollama_port, ollama_models_dir |
| 02-install-vllm.sh | venv_path, vllm_version (optional) |
| 02-install-nginx-tls.sh | app_domain_name, app_upstream_port, cert_country, cert_org |
| 02-configure-log-aggregation.sh | siem_hostname_or_ip, siem_syslog_port, docker_log_driver |
| 02-platform-smoke-test.sh | db_node_hostname, qdrant_port, redis_port, postgres_port |

The database password is prompted by psql and never written to disk by
these scripts. Pick ONE production model server (Ollama or vLLM), not
both.

## Logic flow

Docker → Python venv → Git clone → Postgres → pgvector → databases →
Qdrant → Redis → Ollama and/or vLLM → Nginx/TLS → log aggregation →
smoke test. Container scripts are idempotent (existing containers are
started, not duplicated). Logs go to `/var/log/airag/`.

## Configuration

- Default Postgres version: 16. Default Qdrant ports: 6333/6334.
  Default Redis port: 6379. Default Ollama port: 11434.
- Default databases: airag_docs, airag_standards, airag_audit.
- Default venv path: /opt/airag/venv. Default project dir: /opt/airag.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

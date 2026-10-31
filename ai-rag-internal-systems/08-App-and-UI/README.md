AI-RAG Internal Systems is an on-premises retrieval-augmented generation platform for internal content discovery, question answering over internal documents and regulatory standards (FDA eCFR Title 21, EU GMP, ICH), and document-update recommendations when a standard changes. It is generic: every deployment value is supplied as user input at runtime, and no company-specific value is hardcoded. Every recommendation is a suggestion for human QA and document-owner approval; the system never auto-edits a controlled document.

# 08 — App and UI

## Purpose

The user-facing application and its production services: a FastAPI app
with ask, clause-updates, review-queue, feedback and admin routes, plus
the systemd units that keep the whole stack running across reboots.
Mobile/responsive checks (P8-07) and branding/accessibility work
(P8-08) stay manual.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 08-app-main.py | P8-01, P8-02, P8-03, P8-04, P8-05, P8-06 | FastAPI app: login (session-based, OIDC/LDAP via the SSO config written by 07-configure-sso.sh), home/ask page, clause-updates dashboard, review queue (approve/reject decisions recorded, never applied to documents), feedback endpoint, admin status page. Startup prompts cover every connection value, or set AIRAG_* environment variables to skip prompting. The app NEVER installs packages at runtime: it requires the pinned fastapi==0.115.6, uvicorn==0.30.6 and psycopg[binary]==3.2.3. |
| 08-create-systemd-services.sh | P8-09 | Creates airag-app.service and optionally airag-model.service, enables them, and reminds the operator to reboot-test. Step 1b prompts for the AIRAG_* values listed below and writes them to a mode-600 EnvironmentFile (referenced without a dash prefix, so systemd fails the unit explicitly if the file is missing); both units set StandardInput=null, hardening flags (NoNewPrivileges, PrivateTmp, ProtectSystem=strict, ReadWritePaths for the log directory), UMask 0027, and logs mode 640. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 08-app-main.py | app_host, app_port, app_base_url, qdrant_url, collection_name, model_backend, model_api_url, model_name, embedding_model_name, reranker_model_name, final_top_k, retrieve_top_k (AIRAG_RETRIEVE_TOP_K, default 50), rrf_k (AIRAG_RRF_K, default 60), min_rerank_score, temperature, db_host, db_port, audit_db_name (AIRAG_AUDIT_DB_NAME), standards_db_name (AIRAG_STANDARDS_DB_NAME), docs_db_name (AIRAG_DOCS_DB_NAME), db_user, db_password (secret, via getpass or AIRAG_DB_PASSWORD), prompts_directory, access_map_path, user_groups_map_path (AIRAG_USER_GROUPS_MAP), pipeline_scripts_dir (AIRAG_PIPELINE_DIR), sso_config_path (AIRAG_SSO_CONFIG), session_ttl_seconds (AIRAG_SESSION_TTL, default 3600), session_cookie_secure (AIRAG_SESSION_COOKIE_SECURE), and (LDAP only) ldap_user_dn_template — each also readable from AIRAG_* environment variables |
| 08-create-systemd-services.sh | app_command, app_user, project_dir, log_dir, enable_model_service (yes/no), model_service_command, plus Step 1b: env_file, bind host, bind port, Qdrant URL, collection, model backend, model API URL, model name, audit DB host, audit DB port, audit DB name, audit DB user, audit DB password (read -s), prompts directory, access-map path |

The audit DB password is read with getpass (or from AIRAG_DB_PASSWORD)
and never logged. Login is session-based: OIDC uses the
authorization-code flow with RS256 ID-token verification (JWKS,
issuer/audience/expiry checked before any session is created); LDAP
login verifies the password with a real directory bind. The session
cookie is HttpOnly (samesite lax, Secure when
AIRAG_SESSION_COOKIE_SECURE is yes, max_age = session TTL) and carries
only a random token; identity and groups live server-side. AD groups
come ONLY from the session (ID-token groups claim for OIDC, else the
operator-maintained user-groups map file): there is no ad_groups form
field, no client-supplied groups are accepted anywhere, and a user with
no mapped groups fails closed — /ask abstains without running the
pipeline. /ask is fully wired, not a stub: it runs retrieve →
permission filter → rerank → generate as subprocesses in one temp dir
per request, reranking ONLY the filtered file, with a three-point
abstain rule (empty filtered set, no reranked chunk meeting the minimum
score, or a generated answer with no [doc#chunk] citation — never a
guess). /updates reads the STANDARDS database (clause_changes) and
LEFT-JOINs the suggestions table; when the suggestion step has not been
run yet it prints the clean not-run message ("The suggestion step
has not been run yet, so no suggestions or impact ratings are
available.") instead of crashing, and the DB-table impact ratings
surface here. /review reads the suggestions table and records REAL
document_id and clause_id values read from the suggestion row, with the
session username as reviewer and a decision allowlist of approved /
rejected only. ALL SQL is parameterized (psycopg %s placeholders; no
string interpolation anywhere). EVERY /ask event is audited to
audit_log — answered, abstained, fail-closed no-group, and
pipeline-error events alike — with username, retrieved doc IDs, clause
versions (empty unless a chunk payload supplies them), an answer
excerpt, the abstained flag and the database default timestamp.
Generate (phase 05) itself does no audit logging; the app and the
orchestrator own audit.

## getpass-under-systemd caveat and pipeline stdin design

The app feeds each phase-05 script's prompts on stdin in prompt order
(secrets are stdin lines, never argv). The phase-05 scripts read their
database password with getpass, which falls back to piped stdin when
the process has no controlling terminal (for example under systemd).
Do not run this app attached to an interactive terminal for /ask
requests, or the getpass prompt may read from the terminal instead of
the pipe. Under systemd, run the app only via the Step 1b
EnvironmentFile unit below, with StandardInput=null.

NOTE on the systemd EnvironmentFile: Step 1b prompts for every AIRAG_*
variable the app reads at startup (31 variables, including the LDAP DN
template, which is only used when SSO type is LDAP) and writes them all
to the mode-600 env file, so no manual env-file edit is needed before
enabling the unit. (Earlier versions wrote only a subset and required
the rest to be added by hand; that gap is closed.)

## Logic flow

Start the app (prompts or environment) → user asks → pipeline stages
run in order → answer with citations or abstention → feedback recorded
→ QA reviews clause-update suggestions in the review queue → decisions
recorded in review_decisions. systemd keeps the app (and optionally
the vLLM server) alive across reboots.

## Configuration

- Default bind: 0.0.0.0:8000, behind Nginx TLS (phase 02).
- Default retrieve top-k per path: 50 (AIRAG_RETRIEVE_TOP_K); default
  RRF constant: 60 (AIRAG_RRF_K); default session TTL: 3600 seconds.
- Default collection: airag_docs. Default model backend: ollama.
- Logs: /var/log/airag/app.log (and model.log for vLLM).

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |
| v2 | Fix pass: app fully wired (session login via OIDC/LDAP SSO config with HttpOnly cookie, AD groups server-side only with the ad_groups form field deleted and fail-closed unmapped users, all SQL parameterized with an approved/rejected decision allowlist, /ask retrieve→filter→rerank→generate with filtered-only rerank and three-point abstain, /updates reading the STANDARDS DB with a LEFT JOIN on suggestions and a clean not-run message, /review recording real doc/clause keys, every /ask event audited, pip auto-install removed in favour of pinned fastapi/uvicorn/psycopg, new settings incl. AIRAG_RETRIEVE_TOP_K 50 and AIRAG_RRF_K 60); systemd Step 1b AIRAG_* prompts to a mode-600 EnvironmentFile (no dash prefix, StandardInput=null, hardening flags, UMask 0027, logs 640); getpass-under-systemd caveat and pipeline stdin design documented. |
| v3 | Step 1b extended to prompt for every AIRAG_* variable the app reads at startup (17 added: app base URL, embedding/reranker models, final/retrieve top-k, RRF k, min rerank score, temperature, audit/standards/docs database names, user-groups map, SSO config, pipeline directory, session TTL, session cookie Secure, LDAP DN template); the manual env-file step is removed — the env file is now complete, so the app never prompts under systemd. |

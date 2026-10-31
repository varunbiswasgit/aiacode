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
| 08-app-main.py | P8-01, P8-02, P8-03, P8-04, P8-05, P8-06 | FastAPI app: home/ask page, clause-updates dashboard, review queue (approve/reject decisions recorded, never applied to documents), feedback endpoint, admin status page. Startup prompts cover every connection value, or set AIRAG_* environment variables to skip prompting. |
| 08-create-systemd-services.sh | P8-09 | Creates airag-app.service and optionally airag-model.service, enables them, and reminds the operator to reboot-test. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 08-app-main.py | app_host, app_port, qdrant_url, collection_name, model_backend, model_api_url, model_name, db_host, db_port, db_name, db_user, db_password (secret), prompts_directory, access_map_path — each also readable from AIRAG_* environment variables |
| 08-create-systemd-services.sh | app_command, app_user, project_dir, log_dir, enable_model_service (yes/no), model_service_command |

The audit DB password is read with getpass (or from AIRAG_DB_PASSWORD)
and never logged. The ask route is a wiring stub: it must invoke the
phase-05 scripts in order (hybrid retrieve → permission filter →
rerank → generate) before production use; the permission filter must
run before generation, and review decisions must never modify a
controlled document.

## Logic flow

Start the app (prompts or environment) → user asks → pipeline stages
run in order → answer with citations or abstention → feedback recorded
→ QA reviews clause-update suggestions in the review queue → decisions
recorded in review_decisions. systemd keeps the app (and optionally
the vLLM server) alive across reboots.

## Configuration

- Default bind: 0.0.0.0:8000, behind Nginx TLS (phase 02).
- Default collection: airag_docs. Default model backend: ollama.
- Logs: /var/log/airag/app.log (and model.log for vLLM).

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

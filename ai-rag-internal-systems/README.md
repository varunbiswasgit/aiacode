# AI-RAG Internal Systems

## Purpose

Deployment scripts for an on-prem, internal AI/RAG platform: internal
document Q&A over a small local language model, plus a standards watch
over FDA eCFR Title 21, EU GMP (EudraLex Volume 4) and ICH guidelines
(Q7, Q8, Q9, Q10, Q12). When a standard clause updates, the system
finds the affected internal documents and suggests improvements. A
human (QA or document owner) approves every suggestion; the system
never auto-edits a controlled document.

Development may use Claude as a build assistant, but internal documents
never leave the network at runtime: every data-touching component runs
on premises, and the only outbound traffic is the public standards
fetch through the proxy.

## Project layout

When pushed to the `aiacode` repository, this project lives under
`ai-rag-internal-systems/`, with one self-contained folder per phase.
Each phase folder holds its scripts plus its own README.md and
TESTING.md, following the repository conventions.

| Phase folder | Contents |
|---|---|
| 00-Kickoff-and-governance | Project repo bootstrap |
| 01-Infrastructure | Node identity, patching, AD join, firewall, GPU driver, storage, monitoring, time/logging |
| 02-Platform-software | Docker, Python, Git, Postgres plus pgvector, Qdrant, Redis, Ollama, vLLM, Nginx/TLS, smoke test |
| 03-Internal-ingestion | Source connectors, parsing, metadata, OCR, scheduling, ingest, validation, clause mapping |
| 04-Standards-corpus | eCFR, EU GMP and ICH fetchers, clause parsing, registry, change detection, alerts |
| 05-RAG-core | Embedding, hybrid retrieval, reranking, prompt templates, generation, permission filter, sessions |
| 06-Local-LLM | Model install (Ollama/Hugging Face, quantization), serving, load testing |
| 07-Security | SSO, access map, audit log, PII scan, egress test, vulnerability scan, restore drill |
| 08-App-and-UI | FastAPI application and systemd services |
| 09-Eval-and-pilot | Golden set, evaluation, pilot group, pilot feedback report |
| 10-Rollout | No scripts of its own; reuses phase-03 ingest for phase-2 corpus expansion |

Supporting files in this folder:

- `mapping.csv` — task_id to script_file (filename only, no paths),
  one row per tracker task that has a script.
- The tracker itself (`AI-RAG-IT-Task-List.xlsx`, 117 tasks) is the
  source of truth for scope, order, owners and hour estimates.

## Tracker description

The tracker has one row per task with columns: ID, Phase, Task,
Details/acceptance, Download/source, Run on, Commands, Owner, Priority,
Est. hours, Status, Depends on. 117 tasks across 11 phases, about 709
estimated hours. 80 tasks are covered by the 64 scripts in this
project; the remaining 37 are deliberately script-free (approvals,
policy documents, meetings, manual reviews, and console-only VM work
such as slicing VMs and attaching GPUs).

## Run order

Run the phase folders in numeric order (00 through 10). Inside a
folder, run scripts in tracker task-ID order (the filenames sort the
same way). The dependency chain in the tracker says the same thing:
infrastructure before platform, platform before ingestion and
standards, ingestion plus standards before the RAG core, RAG core
before the app, and evaluation before rollout.

## Script conventions (applies to every script)

- Every variable is prompted interactively at runtime (`read -p` in
  bash with a default shown in brackets, `input()` in Python).
  Required inputs are validated non-empty and re-prompted.
- Secrets (tokens, passwords, webhook URLs) use `read -s` or getpass:
  never echoed, never logged, never written to disk.
- Descriptive variable names only (gpu_server_hostname, model_source,
  huggingface_token). No hardcoded company values: no real IPs,
  domains, tokens or internal URLs; defaults are examples only.
- Bash scripts start with `set -euo pipefail`, print step headers,
  log to a file, fail fast with a readable error, and are idempotent
  where possible (already-installed components are detected and
  skipped).
- Every script header states its purpose, tracker task IDs, prompts
  and prerequisites.
- Model-install scripts always ask: model source (Ollama library vs
  Hugging Face), model name, auth token for gated Hugging Face models,
  and quantization choice (none/AWQ/GPTQ).

## Default topology (prompted, never assumed)

- PRD-AIRAG-GPU01 — inference node (GPU, Ollama or vLLM).
- PRD-AIRAG-APP01 — app/API node (FastAPI behind Nginx TLS).
- PRD-AIRAG-DB01 — data node (PostgreSQL 16 plus pgvector, Qdrant,
  Redis).

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

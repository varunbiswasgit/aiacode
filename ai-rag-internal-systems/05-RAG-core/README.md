AI-RAG Internal Systems is an on-premises retrieval-augmented generation platform for internal content discovery, question answering over internal documents and regulatory standards (FDA eCFR Title 21, EU GMP, ICH), and document-update recommendations when a standard changes. It is generic: every deployment value is supplied as user input at runtime, and no company-specific value is hardcoded. Every recommendation is a suggestion for human QA and document-owner approval; the system never auto-edits a controlled document.

# 05 — RAG core

## Purpose

The retrieval and generation pipeline: embed chunks, hybrid search
(vector plus keyword with reciprocal rank fusion), cross-encoder
reranking, prompt templates, citation-enforced generation with an
abstain rule, AD-group permission filtering, and per-session memory.
These scripts are the pipeline stages; the phase-08 app wires them
together in request order.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 05-embed-documents.py | P5-01, P5-02, P5-03 | Downloads/loads the embedding model (sentence-transformers, installed if missing) and upserts chunk vectors into Qdrant. |
| 05-hybrid-retrieve.py | P5-04 | Embeds the query, runs Qdrant vector search and Postgres full-text search over chunks_fts (built by 03-build-chunks-fts.py), fuses with RRF, writes candidates.json into the prompted output directory. The temporary SQL file is deleted (try/finally). |
| 05-rerank-results.py | P5-05, P5-06 | Downloads/loads the cross-encoder reranker and keeps the top-N candidates with scores. Outputs are chmod 600. |
| 05-create-prompt-templates.py | P5-07 | Writes qa.txt, gap.txt and suggest.txt templates that force citations and forbid invention. |
| 05-generate-answer.py | P5-08, P5-09 | Abstains without calling the model when no chunk clears the score threshold; otherwise calls Ollama/vLLM and rejects answers with no citation. This script has NO audit logging: audit lives in the app (phase 08) and the orchestrator (05-run-qa-pipeline.py). |
| 05-permission-filter.py | P5-10 | Drops candidates the user's AD groups may not see, using the access map (fail closed: unmapped groups get nothing). Must run before reranking/generation. |
| 05-session-store.py | P5-11 | Stores per-session chat turns in Redis with a TTL (set/get/clear). |
| 05-run-qa-pipeline.py | P5-12 | Runs the Q&A pipeline end to end from the command line in the fixed order retrieve → filter → rerank → generate, with filtered-only rerank enforced: the rerank input is forced to the filtered file and is never prompted separately, subset checks verify each step's output against the previous step's, and a missing filtered file (or a filtered path equal to candidates.json) is a hard failure with no fallback to the unfiltered file. An empty filtered list writes an empty reranked.json and generation abstains without calling the model. Every answered AND abstained event is audited (audit insert failure still shows the answer but exits non-zero with an AUDIT FAILED banner). The script takes no command-line arguments; the retrieval database password is prompted by the retrieval script itself via getpass (so the pipeline needs an interactive terminal), and PGPASSWORD is stripped from child processes. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 05-embed-documents.py | chunks_jsonl_path, embedding_model_name (a single model-name prompt; there is no separate source prompt), qdrant_url, collection_name, batch_size, vector_size |
| 05-hybrid-retrieve.py | user_query, qdrant_url, collection_name, embedding_model_name, db_host, db_port, db_name, db_user, db_password (secret), top_k_candidates, rrf_k, candidates_output_dir (any absolute path, default /opt/airag/output) |
| 05-rerank-results.py | user_query, candidates_json_path (default candidates.json, UNCHANGED: when the retrieve output lives elsewhere, enter its full path at this prompt), reranker_model_name, final_top_k, output_json_path |
| 05-create-prompt-templates.py | organisation_name, answer_tone, abstain_message, prompts_directory |
| 05-generate-answer.py | user_question, reranked_json_path, prompts_directory, model_backend, model_api_url, model_name, min_rerank_score, temperature |
| 05-permission-filter.py | username, user_ad_groups, access_map_path, candidates_json_path, output_json_path |
| 05-session-store.py | redis_host, redis_port, session_id, action (set/get/clear), message_text, ttl_seconds |
| 05-run-qa-pipeline.py | user_question, username, user_ad_groups, script_directory, work_directory, access_map_path, filtered_output_path, reranked_output_path, prompts_directory, qdrant_url, collection_name, embedding_model_name, retrieve_db_host, retrieve_db_port, retrieve_db_name, retrieve_db_user, top_k_candidates, rrf_k, reranker_model_name, final_top_k, model_backend, model_api_url, model_name, min_rerank_score, temperature, audit_db_host, audit_db_port, audit_db_name, audit_db_user, audit_db_password (PGPASSWORD or getpass; the retrieval DB password is prompted separately by the retrieval script itself) |

The Postgres password is read with getpass and never saved. The
embedding and reranker models are Hugging Face ids prompted at runtime.

## Logic flow (per question)

1. Hybrid retrieve → candidates.json (top 50 per path, RRF fused).
   Prerequisite: chunks_fts must first be built by
   03-build-chunks-fts.py; retrieval has no placeholder keyword table.
2. Permission filter → candidates-filtered.json (AD groups enforced).
3. Rerank → reranked.json (top 8), fed ONLY the filtered file.
4. Generate → abstain if nothing clears the threshold, else answer with
   citations, rejecting uncited output.
5. Audit → the orchestrator (05-run-qa-pipeline.py) or the app (phase
   08) writes the answered or abstained event to audit_log; the
   generate script itself does no audit logging.

The orchestrator (05-run-qa-pipeline.py) is the CLI runner for this
whole chain: it runs the four stage scripts as subprocesses in the
order above and enforces the filtered-only rerank and the audit step,
so a question cannot skip the filter or the audit.

## Configuration

- Default embedding model: BAAI/bge-large-en-v1.5 (1024 dims).
  sentence-transformers is pinned to 6.2.0: if a different version is
  already loaded, the script installs the pin and exits, asking for a
  re-run so the fresh process loads the pinned version.
- Default reranker: BAAI/bge-reranker-large.
- Default final top-k: 8. Default RRF constant: 60.
- Default temperature: 0.2 for factual answers.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |
| v2 | Fix pass: new orchestrator P5-12 (CLI pipeline runner: retrieve→filter→rerank→generate with filtered-only rerank, hard fail on a missing filtered file, audits answered + abstained events); embed pinned to sentence-transformers 6.2.0 (installs the pin and exits asking for a re-run on a version mismatch) with a single embedding_model_name prompt; hybrid retrieve reads chunks_fts built by 03-build-chunks-fts.py with a new output-directory prompt (any absolute path, default /opt/airag/output) and temp SQL deleted; rerank default input candidates.json UNCHANGED with outputs chmod 600; generate has NO audit logging (audit lives in the app and orchestrator); filter/templates/session: labels only. |

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
| 05-hybrid-retrieve.py | P5-04 | Embeds the query, runs Qdrant vector search and Postgres full-text search, fuses with RRF, writes candidates.json. |
| 05-rerank-results.py | P5-05, P5-06 | Downloads/loads the cross-encoder reranker and keeps the top-N candidates with scores. |
| 05-create-prompt-templates.py | P5-07 | Writes qa.txt, gap.txt and suggest.txt templates that force citations and forbid invention. |
| 05-generate-answer.py | P5-08, P5-09 | Abstains without calling the model when no chunk clears the score threshold; otherwise calls Ollama/vLLM and rejects answers with no citation. |
| 05-permission-filter.py | P5-10 | Drops candidates the user's AD groups may not see, using the access map (fail closed: unmapped groups get nothing). Must run before reranking/generation. |
| 05-session-store.py | P5-11 | Stores per-session chat turns in Redis with a TTL (set/get/clear). |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 05-embed-documents.py | chunks_jsonl_path, embedding_model_name, qdrant_url, collection_name, batch_size, vector_size |
| 05-hybrid-retrieve.py | user_query, qdrant_url, collection_name, embedding_model_name, db_host, db_port, db_name, db_user, db_password (secret), top_k_candidates, rrf_k |
| 05-rerank-results.py | user_query, candidates_json_path, reranker_model_name, final_top_k, output_json_path |
| 05-create-prompt-templates.py | organisation_name, answer_tone, abstain_message, prompts_directory |
| 05-generate-answer.py | user_question, reranked_json_path, prompts_directory, model_backend, model_api_url, model_name, min_rerank_score, temperature |
| 05-permission-filter.py | username, user_ad_groups, access_map_path, candidates_json_path, output_json_path |
| 05-session-store.py | redis_host, redis_port, session_id, action (set/get/clear), message_text, ttl_seconds |

The Postgres password is read with getpass and never saved. The
embedding and reranker models are Hugging Face ids prompted at runtime.

## Logic flow (per question)

1. Hybrid retrieve → candidates.json (top 50 per path, RRF fused).
2. Permission filter → candidates-filtered.json (AD groups enforced).
3. Rerank → reranked.json (top 8).
4. Generate → abstain if nothing clears the threshold, else answer with
   citations, rejecting uncited output.

## Configuration

- Default embedding model: BAAI/bge-large-en-v1.5 (1024 dims).
- Default reranker: BAAI/bge-reranker-large.
- Default final top-k: 8. Default RRF constant: 60.
- Default temperature: 0.2 for factual answers.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

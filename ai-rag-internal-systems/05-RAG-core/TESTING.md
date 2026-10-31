# Testing — 05 RAG core

Manual test cases. Run each case, then tick the Pass column. Use a
small chunk JSONL (20-50 chunks) and a test Qdrant collection.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Chunk JSONL, Qdrant running, test collection name | Run 05-embed-documents.py with a small test model | Collection created, point count equals chunk count, vector size matches the model. NOTE: if a different sentence-transformers version is loaded, the script installs the pinned 6.2.0 and exits asking for a re-run — run it a second time so the fresh process loads the pinned version | ☐ |
| Test Qdrant collection plus a chunks_fts table built by 03-build-chunks-fts.py | Run 05-hybrid-retrieve.py with a query about the sample corpus, answering the candidates-output-directory prompt | candidates.json written into the prompted output directory, results include hits from both vector and keyword paths; the temporary SQL file is deleted | ☐ |
| candidates.json from the retrieve test | Run 05-rerank-results.py with final top-k 3 | reranked.json has at most 3 rows, each with a rerank_score, sorted descending | ☐ |
| Empty prompts directory | Run 05-create-prompt-templates.py with test organisation and tone | qa.txt, gap.txt, suggest.txt exist, each requires citations and forbids invention | ☐ |
| reranked.json plus a running test model server | Run 05-generate-answer.py with a question the corpus answers | Answer printed with at least one doc citation in brackets | ☐ |
| reranked.json with no chunk above the threshold | Run 05-generate-answer.py with a high minimum score | Script prints the abstain message and never calls the model | ☐ |
| candidates.json plus a test access map | Run 05-permission-filter.py for a user whose groups allow only one doc prefix | Output keeps only chunks from allowed docs, unmapped groups keep nothing | ☐ |
| Redis running on the test node | Run 05-session-store.py set, then get, then clear for a test session | Get returns the stored message, clear empties the session | ☐ |
| Test Qdrant collection, chunks_fts table, test access map and a running test model server | Run 05-run-qa-pipeline.py with a test question, user and groups | Answer printed with citations (or the explicit not-found response when nothing supports an answer), reranking used only the filtered candidates, and an audit_log row records the answered or abstained event | ☐ |

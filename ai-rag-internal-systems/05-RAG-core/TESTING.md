# Testing — 05 RAG core

Manual test cases. Run each case, then tick the Pass column. Use a
small chunk JSONL (20-50 chunks) and a test Qdrant collection.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Chunk JSONL, Qdrant running, test collection name | Run 05-embed-documents.py with a small test model | Collection created, point count equals chunk count, vector size matches the model | ☐ |
| Test Qdrant collection plus a chunks_fts table | Run 05-hybrid-retrieve.py with a query about the sample corpus | candidates.json written, results include hits from both vector and keyword paths | ☐ |
| candidates.json from the retrieve test | Run 05-rerank-results.py with final top-k 3 | reranked.json has at most 3 rows, each with a rerank_score, sorted descending | ☐ |
| Empty prompts directory | Run 05-create-prompt-templates.py with test organisation and tone | qa.txt, gap.txt, suggest.txt exist, each requires citations and forbids invention | ☐ |
| reranked.json plus a running test model server | Run 05-generate-answer.py with a question the corpus answers | Answer printed with at least one doc citation in brackets | ☐ |
| reranked.json with no chunk above the threshold | Run 05-generate-answer.py with a high minimum score | Script prints the abstain message and never calls the model | ☐ |
| candidates.json plus a test access map | Run 05-permission-filter.py for a user whose groups allow only one doc prefix | Output keeps only chunks from allowed docs, unmapped groups keep nothing | ☐ |
| Redis running on the test node | Run 05-session-store.py set, then get, then clear for a test session | Get returns the stored message, clear empties the session | ☐ |

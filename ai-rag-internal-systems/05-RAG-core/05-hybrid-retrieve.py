#!/usr/bin/env python3
"""Hybrid retrieval: vector search (Qdrant) + keyword search (Postgres FTS),
fused with reciprocal rank fusion (RRF).

Tracker task: P5-04 (hybrid retriever)
Prompts: user_query, qdrant_url, collection_name, embedding_model_name,
         db_host, db_port, db_name, db_user, db_password (secret),
         top_k_candidates, rrf_k, candidates_output_dir
Prerequisites: Qdrant collection from 05-embed-documents.py; a Postgres
  table named chunks_fts (doc_id, chunk_index, text, tsv), built by
  03-Internal-ingestion/03-build-chunks-fts.py from the chunk JSONL
  produced by 03-ingest-corpus.py. sentence-transformers is installed
  automatically if missing.
Output: ranked candidates JSON printed and saved to
  <candidates_output_dir>/candidates.json.
"""
import getpass
import json
import os
import subprocess
import sys
import tempfile
import urllib.request


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_int(label, default):
    while True:
        raw = input(f"{label} [{default}]: ").strip() or str(default)
        if raw.isdigit() and int(raw) > 0:
            return int(raw)
        print("Enter a positive whole number.")


def prompt_secret(label):
    while True:
        value = getpass.getpass(f"{label}: ").strip()
        if value:
            return value
        print("A value is required.")


def qdrant_search(url, collection, vector, limit):
    payload = {"vector": vector, "limit": limit, "with_payload": True}
    request = urllib.request.Request(
        f"{url}/collections/{collection}/points/search",
        data=json.dumps(payload).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(request, timeout=60) as response:
        return json.loads(response.read().decode("utf-8"))["result"]


def prompt_output_dir():
    # Any absolute path is accepted, not just paths under /opt/airag:
    # the web app passes a per-request temp dir and the QA pipeline
    # passes its work directory, so restricting output to /opt/airag
    # would break both callers. The path must still be absolute.
    while True:
        value = prompt("Candidates output directory", "/opt/airag/output")
        path = os.path.normpath(value)
        if not os.path.isabs(path):
            print("Enter an absolute path.")
            continue
        return path


def postgres_search(host, port, db, user, password, query, limit):
    sql = ("SELECT doc_id, chunk_index, text FROM chunks_fts "
           "WHERE tsv @@ plainto_tsquery('english', $QUERY$) "
           "ORDER BY ts_rank(tsv, plainto_tsquery('english', $QUERY$)) DESC "
           f"LIMIT {limit};").replace("$QUERY$", "'" + query.replace("'", "''") + "'")
    with tempfile.NamedTemporaryFile("w", suffix=".sql", delete=False) as handle:
        handle.write(sql)
        sql_path = handle.name
    env = {"PGPASSWORD": password, "PATH": "/usr/bin:/bin:/usr/local/bin"}
    try:
        result = subprocess.run(
            ["psql", "-h", host, "-p", port, "-U", user, "-d", db,
             "-t", "-A", "-F", "\t", "-f", sql_path],
            capture_output=True, text=True, env=env)
    finally:
        try:
            os.unlink(sql_path)
        except OSError:
            pass
    if result.returncode != 0:
        print(f"WARNING: Postgres FTS failed ({result.stderr.strip()}); continuing vector-only.")
        return []
    rows = []
    for line in result.stdout.strip().splitlines():
        if not line.strip():
            continue
        doc_id, chunk_index, text = line.split("\t", 2)
        rows.append({"doc_id": doc_id, "chunk_index": int(chunk_index), "text": text})
    return rows


def main():
    print("=== Step 1: Retrieval settings ===")
    user_query = prompt("User query")
    qdrant_url = prompt("Qdrant base URL", "http://PRD-AIRAG-DB01:6333")
    collection_name = prompt("Qdrant collection", "airag_docs")
    embedding_model_name = prompt("Embedding model (same as embedding step)", "BAAI/bge-large-en-v1.5")
    db_host = prompt("Postgres host", "PRD-AIRAG-DB01")
    db_port = prompt("Postgres port", "5432")
    db_name = prompt("Postgres database", "airag_docs")
    db_user = prompt("Postgres user", "airag_app")
    db_password = prompt_secret("Postgres password (not echoed, never saved)")
    top_k_candidates = prompt_int("Candidates per path", 50)
    rrf_k = prompt_int("RRF constant k", 60)
    candidates_output_dir = prompt_output_dir()
    os.makedirs(candidates_output_dir, exist_ok=True)

    print("=== Step 2: Searching ===")
    try:
        from sentence_transformers import SentenceTransformer
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "sentence-transformers"], check=True)
        from sentence_transformers import SentenceTransformer
    model = SentenceTransformer(embedding_model_name)
    query_vector = model.encode([user_query], normalize_embeddings=True).tolist()[0]
    vector_hits = qdrant_search(qdrant_url, collection_name, query_vector, top_k_candidates)
    keyword_hits = postgres_search(db_host, db_port, db_name, db_user, db_password,
                                   user_query, top_k_candidates)

    print("=== Step 3: RRF fusion ===")
    scores = {}
    for rank, hit in enumerate(vector_hits, 1):
        key = (hit["payload"]["doc_id"], hit["payload"]["chunk_index"])
        scores.setdefault(key, {"doc_id": key[0], "chunk_index": key[1],
                                "text": hit["payload"]["text"], "score": 0.0})
        scores[key]["score"] += 1.0 / (rrf_k + rank)
    for rank, hit in enumerate(keyword_hits, 1):
        key = (hit["doc_id"], hit["chunk_index"])
        scores.setdefault(key, {"doc_id": key[0], "chunk_index": key[1],
                                "text": hit["text"], "score": 0.0})
        scores[key]["score"] += 1.0 / (rrf_k + rank)
    ranked = sorted(scores.values(), key=lambda item: item["score"], reverse=True)
    candidates_path = os.path.join(candidates_output_dir, "candidates.json")
    with open(candidates_path, "w", encoding="utf-8") as handle:
        json.dump(ranked, handle, indent=2)
    print(f"Done. {len(ranked)} fused candidates -> {candidates_path}")
    for item in ranked[:10]:
        print(f"  [{item['score']:.4f}] {item['doc_id']}#{item['chunk_index']}: {item['text'][:120]}")


if __name__ == "__main__":
    main()

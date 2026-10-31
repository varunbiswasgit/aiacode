#!/usr/bin/env python3
"""Embed chunk JSONL and upsert the vectors into Qdrant.

Tracker tasks: P5-01 (choose + download embedding model),
               P5-02 (embedding service), P5-03 (embed corpora)
Prompts: chunks_jsonl_path, embedding_model_name, embedding_model_source
         (huggingface/sentence-transformers id), qdrant_url, collection_name,
         batch_size, vector_size
Prerequisites: Python 3.12, Qdrant running. sentence-transformers is
  installed automatically if missing (pip install sentence-transformers).
  The vector size must match the model (BGE-large = 1024, E5-large = 1024,
  nomic-embed-text = 768); the script verifies it on the first batch.
"""
import json
import subprocess
import sys
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


def qdrant_request(url, method, payload=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    request = urllib.request.Request(url, data=data, method=method,
                                     headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(request, timeout=120) as response:
        return json.loads(response.read().decode("utf-8"))


def ensure_model():
    try:
        from sentence_transformers import SentenceTransformer
        return SentenceTransformer
    except ImportError:
        print("sentence-transformers not found; installing...")
        subprocess.run([sys.executable, "-m", "pip", "install", "sentence-transformers"],
                       check=True)
        from sentence_transformers import SentenceTransformer
        return SentenceTransformer


def main():
    print("=== Step 1: Embedding settings ===")
    chunks_jsonl_path = prompt("Chunks JSONL", "/data/staging/chunks-pilot.jsonl")
    embedding_model_name = prompt("Embedding model (Hugging Face id)", "BAAI/bge-large-en-v1.5")
    qdrant_url = prompt("Qdrant base URL", "http://PRD-AIRAG-DB01:6333")
    collection_name = prompt("Qdrant collection", "airag_docs")
    batch_size = prompt_int("Batch size", 32)
    vector_size = prompt_int("Vector size for this model", 1024)

    print("=== Step 2: Loading model ===")
    SentenceTransformer = ensure_model()
    model = SentenceTransformer(embedding_model_name)

    print("=== Step 3: Ensuring collection ===")
    collections = qdrant_request(f"{qdrant_url}/collections", "GET")
    existing = [item["name"] for item in collections.get("result", {}).get("collections", [])]
    if collection_name not in existing:
        qdrant_request(f"{qdrant_url}/collections/{collection_name}", "PUT",
                       {"vectors_config": {"size": vector_size, "distance": "Cosine"}})
        print(f"Created collection {collection_name}.")
    else:
        print(f"Collection {collection_name} already exists.")

    print("=== Step 4: Embedding and upserting ===")
    with open(chunks_jsonl_path, encoding="utf-8") as handle:
        chunks = [json.loads(line) for line in handle]
    total = 0
    for start in range(0, len(chunks), batch_size):
        batch = chunks[start:start + batch_size]
        vectors = model.encode([chunk["text"] for chunk in batch],
                               normalize_embeddings=True).tolist()
        if len(vectors[0]) != vector_size:
            print(f"ERROR: model produced {len(vectors[0])} dims, expected {vector_size}.")
            print("Re-run and enter the correct vector size.")
            raise SystemExit(1)
        points = [{
            "id": start + index,
            "vector": vector,
            "payload": {"doc_id": chunk.get("doc_id"), "chunk_index": chunk.get("chunk_index"),
                        "source_path": chunk.get("source_path"), "text": chunk["text"]},
        } for index, (chunk, vector) in enumerate(zip(batch, vectors))]
        qdrant_request(f"{qdrant_url}/collections/{collection_name}/points?wait=true",
                       "PUT", {"points": points})
        total += len(batch)
        print(f"  Upserted {total}/{len(chunks)}")
    print(f"Done. {total} vectors in collection {collection_name}.")


if __name__ == "__main__":
    main()

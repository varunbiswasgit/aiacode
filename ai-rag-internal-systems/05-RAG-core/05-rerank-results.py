#!/usr/bin/env python3
"""Rerank retrieval candidates with a cross-encoder and keep the top N.

Tracker tasks: P5-05 (choose + download reranker), P5-06 (rerank step)
Prompts: user_query, candidates_json_path, reranker_model_name,
         final_top_k, output_json_path
Prerequisites: Python 3.12; candidates.json from 05-hybrid-retrieve.py.
  sentence-transformers is installed automatically if missing.
  The reranker model is downloaded on first use (Hugging Face).
Output: top-N reranked chunks JSON with scores.
"""
import json
import subprocess
import sys


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


def main():
    print("=== Step 1: Rerank settings ===")
    user_query = prompt("User query")
    candidates_json_path = prompt("Candidates JSON", "candidates.json")
    reranker_model_name = prompt("Reranker model (Hugging Face id)", "BAAI/bge-reranker-large")
    final_top_k = prompt_int("Final top-k to keep", 8)
    output_json_path = prompt("Output JSON path", "reranked.json")

    print("=== Step 2: Loading reranker (downloads on first use) ===")
    try:
        from sentence_transformers import CrossEncoder
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "sentence-transformers"], check=True)
        from sentence_transformers import CrossEncoder
    reranker = CrossEncoder(reranker_model_name)

    with open(candidates_json_path, encoding="utf-8") as handle:
        candidates = json.load(handle)
    pairs = [(user_query, candidate["text"]) for candidate in candidates]
    scores = reranker.predict(pairs).tolist()
    for candidate, score in zip(candidates, scores):
        candidate["rerank_score"] = float(score)
    ranked = sorted(candidates, key=lambda item: item["rerank_score"], reverse=True)[:final_top_k]
    with open(output_json_path, "w", encoding="utf-8") as handle:
        json.dump(ranked, handle, indent=2)
    print(f"Done. Top {len(ranked)} -> {output_json_path}")
    for item in ranked:
        print(f"  [{item['rerank_score']:.4f}] {item['doc_id']}#{item['chunk_index']}")


if __name__ == "__main__":
    main()

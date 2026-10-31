#!/usr/bin/env python3
"""Generate an answer from reranked chunks via the local model, with
citation enforcement and an abstain rule.

Tracker tasks: P5-08 (citation rule), P5-09 (abstain rule)
Prompts: user_question, reranked_json_path, prompts_directory,
         model_backend (ollama/vllm), model_api_url, model_name,
         min_rerank_score, temperature
Prerequisites: reranked.json from 05-rerank-results.py; a running local
  model (Ollama on :11434 or vLLM OpenAI-compatible on :8000).
How it works: if no chunk scores above min_rerank_score, it abstains
  WITHOUT calling the model. Otherwise it calls the model with the QA
  template and rejects the answer if it contains no [doc#chunk] citation.
"""
import json
import re
import urllib.request

CITATION_PATTERN = re.compile(r"\[[^\]]+#\d+\]")


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_float(label, default):
    while True:
        raw = input(f"{label} [{default}]: ").strip() or str(default)
        try:
            return float(raw)
        except ValueError:
            print("Enter a number.")


def call_model(backend, api_url, model_name, prompt_text, temperature):
    if backend == "ollama":
        payload = {"model": model_name, "prompt": prompt_text, "stream": False,
                   "options": {"temperature": temperature}}
        endpoint = f"{api_url.rstrip('/')}/api/generate"
        key = "response"
    else:
        payload = {"model": model_name,
                   "messages": [{"role": "user", "content": prompt_text}],
                   "temperature": temperature}
        endpoint = f"{api_url.rstrip('/')}/v1/chat/completions"
        key = None
    request = urllib.request.Request(
        endpoint, data=json.dumps(payload).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(request, timeout=300) as response:
        data = json.loads(response.read().decode("utf-8"))
    if key:
        return data[key]
    return data["choices"][0]["message"]["content"]


def main():
    print("=== Step 1: Generation settings ===")
    user_question = prompt("User question")
    reranked_json_path = prompt("Reranked chunks JSON (example — replace with your value)", "reranked.json")
    prompts_directory = prompt("Prompts directory (example — replace with your value)", "/opt/airag/prompts")
    model_backend = prompt("Model backend (ollama/vllm) (example — replace with your value)", "ollama").lower()
    model_api_url = prompt("Model API URL (example — replace with your value)", "http://PRD-AIRAG-GPU01:11434"
                           if model_backend == "ollama" else "http://PRD-AIRAG-GPU01:8000")
    model_name = prompt("Model name (example — replace with your value)", "llama3.1:8b")
    min_rerank_score = prompt_float("Minimum rerank score to answer (example — replace with your value)", 0.0)
    temperature = prompt_float("Temperature (0.1-0.3 for factual answers) (example — replace with your value)", 0.2)

    with open(reranked_json_path, encoding="utf-8") as handle:
        chunks = json.load(handle)
    usable = [chunk for chunk in chunks
              if chunk.get("rerank_score", 0) >= min_rerank_score]
    if not usable:
        print("Not found in internal documents.")
        print("(Abstained: no chunk scored above the threshold - no guess was made.)")
        return

    with open(f"{prompts_directory}/qa.txt", encoding="utf-8") as handle:
        template = handle.read()
    context = "\n\n".join(
        f"[{chunk['doc_id']}#{chunk['chunk_index']}] {chunk['text']}" for chunk in usable)
    prompt_text = template.replace("{context}", context).replace("{question}", user_question)

    print("=== Step 2: Calling the local model ===")
    answer = call_model(model_backend, model_api_url, model_name, prompt_text, temperature)

    print("=== Step 3: Citation check ===")
    if not CITATION_PATTERN.search(answer):
        print("WARNING: model answer contained no citation - rejecting it.")
        print("Not found in internal documents.")
        return
    print("\n--- Answer ---")
    print(answer)


if __name__ == "__main__":
    main()

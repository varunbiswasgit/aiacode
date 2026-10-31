#!/usr/bin/env python3
"""Load-test the local model server: latency percentiles and throughput.

Tracker task: P6-05 (load-test the model server)
Prompts: model_backend (ollama/vllm), model_api_url, model_name,
         concurrent_users, prompts_per_user, prompt_text, output_report_path
Prerequisites: model server running (06-serve-language-model.sh).
  Uses only the standard library (urllib, threading, statistics).
Output: a report with p50/p95 latency, tokens/sec estimate and error count.
"""
import json
import statistics
import threading
import time
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


def call_once(backend, api_url, model_name, prompt_text):
    if backend == "ollama":
        payload = {"model": model_name, "prompt": prompt_text, "stream": False}
        endpoint = f"{api_url.rstrip('/')}/api/generate"
    else:
        payload = {"model": model_name,
                   "messages": [{"role": "user", "content": prompt_text}],
                   "max_tokens": 128}
        endpoint = f"{api_url.rstrip('/')}/v1/chat/completions"
    request = urllib.request.Request(
        endpoint, data=json.dumps(payload).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json"})
    start = time.perf_counter()
    with urllib.request.urlopen(request, timeout=300) as response:
        data = json.loads(response.read().decode("utf-8"))
    elapsed = time.perf_counter() - start
    if backend == "ollama":
        text = data.get("response", "")
    else:
        text = data["choices"][0]["message"]["content"]
    return elapsed, len(text.split())


def main():
    print("=== Step 1: Load-test settings ===")
    model_backend = prompt("Model backend (ollama/vllm) (example — replace with your value)", "ollama").lower()
    model_api_url = prompt("Model API URL (example — replace with your value)", "http://PRD-AIRAG-GPU01:11434"
                           if model_backend == "ollama" else "http://PRD-AIRAG-GPU01:8000")
    model_name = prompt("Model name (example — replace with your value)", "llama3.1:8b")
    concurrent_users = prompt_int("Concurrent users (example — replace with your value)", 10)
    prompts_per_user = prompt_int("Prompts per user (example — replace with your value)", 5)
    prompt_text = prompt("Test prompt (example — replace with your value)", "Summarise good manufacturing practice in two sentences.")
    output_report_path = prompt("Report path (example — replace with your value)", "/tmp/airag-loadtest-report.md")

    print("=== Step 2: Running load test ===")
    latencies = []
    total_tokens = 0
    errors = 0
    lock = threading.Lock()

    def worker():
        nonlocal total_tokens, errors
        for _ in range(prompts_per_user):
            try:
                elapsed, tokens = call_once(model_backend, model_api_url, model_name, prompt_text)
                with lock:
                    latencies.append(elapsed)
                    total_tokens += tokens
            except Exception as error:
                with lock:
                    errors += 1
                print(f"  request failed: {error}")

    threads = [threading.Thread(target=worker) for _ in range(concurrent_users)]
    wall_start = time.perf_counter()
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    wall_time = time.perf_counter() - wall_start

    total_requests = concurrent_users * prompts_per_user
    successful = len(latencies)
    p50 = statistics.median(latencies) if latencies else 0
    p95 = sorted(latencies)[int(0.95 * len(latencies)) - 1] if latencies else 0
    report = (
        f"# AI-RAG model load test\n\n"
        f"- Backend: {model_backend}, model: {model_name}\n"
        f"- Users: {concurrent_users}, prompts/user: {prompts_per_user}, "
        f"total requests: {total_requests}\n"
        f"- Successful: {successful}, errors: {errors}\n"
        f"- p50 latency: {p50:.2f}s, p95 latency: {p95:.2f}s\n"
        f"- Approx throughput: {total_tokens / wall_time:.1f} tokens/sec over {wall_time:.1f}s\n"
    )
    with open(output_report_path, "w", encoding="utf-8") as handle:
        handle.write(report)
    print(report)
    print(f"Report saved to {output_report_path}. Also watch nvidia-smi during the test.")
    print("Record these numbers in /docs/perf.md (P6-05 acceptance).")


if __name__ == "__main__":
    main()

# 06 — Local LLM

## Purpose

Install, serve, load-test and configure the small local language model
on the GPU node. Model choice, source and quantization are all prompted
at runtime; nothing about the model is hardcoded beyond the example
defaults shown in brackets.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 06-install-language-model.sh | P6-02, P6-03 | Pulls an Ollama model, or downloads a Hugging Face model (with optional token) and records the quantization choice (none/AWQ/GPTQ). |
| 06-serve-language-model.sh | P6-04, P6-06, P6-07 | Writes the generation config, starts Ollama or vLLM serving, and provides a healthcheck action used as the GPU-outage fallback probe. |
| 06-loadtest-model.py | P6-05 | Fires concurrent prompts at the model server and reports p50/p95 latency, tokens/sec and errors. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 06-install-language-model.sh | model_source (ollama/huggingface), model_name, huggingface_token (secret, HF only), quantization (none/AWQ/GPTQ), ollama_port, models_dir |
| 06-serve-language-model.sh | action (serve/healthcheck), serve_backend (ollama/vllm), model_name, serve_port, gpu_host, context_length, temperature, top_p, max_tokens, config_path, venv_path |
| 06-loadtest-model.py | model_backend, model_api_url, model_name, concurrent_users, prompts_per_user, prompt_text, output_report_path |

The Hugging Face token is read with `read -s`, used only for the
download, then unset; it is never echoed, logged or written to disk.
For gated models (for example Meta Llama) the user must first accept
the licence on Hugging Face.

## Logic flow

Install (choose source → download → smoke test) → serve (write config
→ start server → verify endpoint) → load-test (concurrent prompts →
report) → record the numbers in the perf doc. The healthcheck action
returns non-zero when the model server is down, which is the signal the
app uses to queue requests and show a status banner instead of failing
silently.

## Configuration

- Default models: llama3.1:8b (Ollama) or meta-llama/Llama-3.1-8B-Instruct (HF).
- Default ports: Ollama 11434, vLLM 8000.
- Default generation: temperature 0.2, top_p 0.9, max_tokens 1024,
  context length 8192.
- Model selection and licence approval (P6-01) stay manual decisions.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

AI-RAG Internal Systems is an on-premises retrieval-augmented generation platform for internal content discovery, question answering over internal documents and regulatory standards (FDA eCFR Title 21, EU GMP, ICH), and document-update recommendations when a standard changes. It is generic: every deployment value is supplied as user input at runtime, and no company-specific value is hardcoded. Every recommendation is a suggestion for human QA and document-owner approval; the system never auto-edits a controlled document.

# 06 — Local LLM

## Purpose

Install, serve, load-test and configure the small local language model
on the GPU node. Model choice, source and quantization are all prompted
at runtime; nothing about the model is hardcoded beyond the example
defaults shown in brackets.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 06-install-language-model.sh | P6-02, P6-03 | Pulls an Ollama model, or downloads a Hugging Face model (with optional token) and records the quantization choice (none/AWQ/GPTQ). Checksum is ENFORCED: the Hugging Face branch requires a revision and a SHA256 manifest and hard-fails on any mismatch, missing file, missing manifest or empty manifest (do not serve that download); the Ollama branch relies on pull digest verification (each blob SHA256 checked against the registry manifest, hard fail on mismatch) and smoke-tests with ollama run. Operator values reach the download code via sys.argv into a quoted heredoc, never interpolated into python -c source. huggingface_hub is pinned to 0.34.0. |
| 06-serve-language-model.sh | P6-04, P6-06, P6-07 | Writes the generation config, starts Ollama or vLLM serving, and provides a healthcheck action used as the GPU-outage fallback probe. For vLLM, an existing-process/port guard exits 0 with "already running" instead of starting a second server, and serving binds an explicit internal --host from a prompted bind-IP prompt (0.0.0.0 rejected). gpu_host is now required (no default): it must be typed explicitly. |
| 06-loadtest-model.py | P6-05 | Fires concurrent prompts at the model server and reports p50/p95 latency, tokens/sec and errors. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 06-install-language-model.sh | model_source (ollama/huggingface), model_name, huggingface_token (secret, HF only), quantization (none/AWQ/GPTQ), ollama_port, models_dir, model_revision (HF only, full commit hash from the model card; a non-40-hex value is accepted only with a warning, and then the checksum check is the only integrity gate), model_checksums_path (HF only, SHA256 manifest in "<sha256>  <relative-path>" format from the model card/repo) |
| 06-serve-language-model.sh | action (serve/healthcheck), serve_backend (ollama/vllm), model_name, serve_port, gpu_host (required, no default), context_length, temperature, top_p, max_tokens, config_path, venv_path, serve_bind_ip (vLLM serve only, internal interface IP or 127.0.0.1, never 0.0.0.0) |
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
| v2 | Fix pass: model install checksum ENFORCED (Hugging Face revision + SHA256 manifest required, hard fail on mismatch; Ollama branch relies on pull digest verification), values passed via argv never interpolated into python -c, huggingface_hub pinned to 0.34.0, smoke test via ollama run; serve: vLLM existing-process/port guard exits 0 "already running", explicit internal --host prompt, gpu_host now required (no default); loadtest: labels only. |

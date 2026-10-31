# Testing — 06 Local LLM

Manual test cases. Run each case, then tick the Pass column. Use a
small test model (for example llama3.1:8b) on the GPU test node.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| GPU test node, Ollama installed | Run 06-install-language-model.sh with source ollama and a small test model | ollama list shows the model, smoke-test generate returns text | ☐ |
| GPU test node, Ollama installed | Run 06-install-language-model.sh with source huggingface, a public small model, quantization none | Model files appear in the chosen models directory, no token echoed in output | ☐ |
| GPU test node, model installed | Run 06-serve-language-model.sh with action serve, then again with action healthcheck | Serve starts the endpoint, healthcheck prints HEALTHY and exits 0 | ☐ |
| GPU test node, model serving, server stopped | Run 06-serve-language-model.sh with action healthcheck | Healthcheck prints UNHEALTHY and exits non-zero (the fallback signal) | ☐ |
| GPU test node, model serving | Run 06-loadtest-model.py with 2 users and 2 prompts each | Report written with p50, p95, tokens per second and error count | ☐ |

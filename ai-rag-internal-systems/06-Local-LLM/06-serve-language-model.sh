#!/usr/bin/env bash
# =============================================================================
# Purpose:      Serve the local model (Ollama or vLLM), write the generation
#               config, and provide a health check / fallback probe.
# Tracker tasks: P6-04 (configure model serving), P6-06 (generation params),
#               P6-07 (fallback for GPU node outage)
# Prompts:      serve_backend (ollama/vllm), model_name, serve_port,
#               context_length, temperature, top_p, max_tokens,
#               config_path, action (serve/healthcheck)
# Prerequisites: Model installed (06-install-language-model.sh). Run ON GPU
#               node for serve; the healthcheck action can run from the app
#               node against the GPU node's URL.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-serve-model.log"
sudo mkdir -p /var/log/airag
exec > >(sudo tee -a "${LOG_FILE}" >/dev/null) 2>&1 || true

prompt() { # prompt <label> <default>
    local value
    while true; do
        read -rp "$1 [$2]: " value
        value="${value:-$2}"
        if [ -n "${value}" ]; then echo "${value}"; return; fi
        echo "A value is required. Please try again."
    done
}
prompt_float() {
    local value
    while true; do
        read -rp "$1 [$2]: " value
        value="${value:-$2}"
        if [[ "${value}" =~ ^[0-9]+(\.[0-9]+)?$ ]]; then echo "${value}"; return; fi
        echo "Enter a number."
    done
}

echo "=== Step 1: Serving settings ==="
action="$(prompt "Action (serve/healthcheck)" "serve")"
serve_backend="$(prompt "Backend (ollama/vllm)" "ollama")"
model_name="$(prompt "Model name" "llama3.1:8b")"
serve_port="$(prompt "Serve port" "11434")"
gpu_host="$(prompt "GPU node hostname (for healthcheck from elsewhere)" "PRD-AIRAG-GPU01")"

if [ "${action}" = "healthcheck" ]; then
    echo "=== Health check ==="
    if curl -fsS "http://${gpu_host}:${serve_port}/api/tags" > /dev/null 2>&1 || \
       curl -fsS "http://${gpu_host}:${serve_port}/v1/models" > /dev/null 2>&1; then
        echo "HEALTHY: model server responding on ${gpu_host}:${serve_port}."
    else
        echo "UNHEALTHY: model server NOT responding."
        echo "Fallback behaviour: the app must queue requests and show a status banner (P6-07)."
        exit 1
    fi
    exit 0
fi

context_length="$(prompt "Context length (tokens)" "8192")"
temperature="$(prompt_float "Temperature (0.1-0.3 for factual Q&A)" "0.2")"
top_p="$(prompt_float "Top-p" "0.9")"
max_tokens="$(prompt "Max output tokens" "1024")"
config_path="$(prompt "Generation config path" "/opt/airag/config.yaml")"
venv_path="$(prompt "Venv path (vLLM only, blank for Ollama)" "")"

echo "=== Step 2: Write generation config ==="
cat > "${config_path}" <<EOF
# AI-RAG generation parameters (P6-06)
model_name: ${model_name}
backend: ${serve_backend}
temperature: ${temperature}
top_p: ${top_p}
max_tokens: ${max_tokens}
context_length: ${context_length}
EOF
echo "Config written to ${config_path}."

echo "=== Step 3: Serve ==="
if [ "${serve_backend}" = "ollama" ]; then
    sudo systemctl enable --now ollama
    curl -fsS "http://localhost:${serve_port}/api/tags" | head -c 300; echo
    echo "Ollama serving. Test: ollama run ${model_name} 'Say OK.'"
else
    if [ -z "${venv_path}" ]; then echo "ERROR: venv path required for vLLM."; exit 1; fi
    nohup "${venv_path}/bin/python" -m vllm.entrypoints.openai.api_server \
        --model "${model_name}" --port "${serve_port}" \
        --max-model-len "${context_length}" \
        > /var/log/airag/vllm.log 2>&1 &
    sleep 15
    curl -fsS "http://localhost:${serve_port}/v1/models" | head -c 300; echo
    echo "vLLM serving (log: /var/log/airag/vllm.log)."
fi
echo "Done."

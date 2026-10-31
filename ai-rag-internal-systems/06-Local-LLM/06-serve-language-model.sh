#!/usr/bin/env bash
# =============================================================================
# Purpose:      Serve the local model (Ollama or vLLM), write the generation
#               config, and provide a health check / fallback probe.
# Tracker tasks: P6-04 (configure model serving), P6-06 (generation params),
#               P6-07 (fallback for GPU node outage)
# Prompts:      serve_backend (ollama/vllm), model_name, serve_port,
#               gpu_host, serve_bind_ip (vLLM serve only),
#               context_length, temperature, top_p, max_tokens,
#               config_path, venv_path, action (serve/healthcheck)
# Prerequisites: Model installed (06-install-language-model.sh). Run ON GPU
#               node for serve; the healthcheck action can run from the app
#               node against the GPU node's URL.
# Security:
#   - vLLM is bound to an explicit internal --host (prompted below),
#     never to all interfaces. The vLLM OpenAI server has no
#     authentication by itself, and its --api-key option does not gate
#     every endpoint, so binding alone is not authentication: keep the
#     port inside the internal network and front it with an
#     authenticated TLS proxy before any non-local client uses it.
#   - The Ollama branch is left as-is: `systemctl enable --now` is
#     already idempotent, so it needs no duplicate-server guard. The
#     vLLM branch, which launches with nohup, gets the guard below.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-serve-model.log"
sudo mkdir -p /var/log/airag
sudo touch "${LOG_FILE}"
sudo chmod 640 "${LOG_FILE}"
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
prompt_bind_ip() { # prompt_bind_ip <label>
    local value
    while true; do
        read -rp "$1: " value
        if [[ "${value}" =~ ^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+$ ]] && [ "${value}" != "0.0.0.0" ]; then
            echo "${value}"; return
        fi
        echo "Enter an IPv4 address for the internal interface (or 127.0.0.1 for localhost-only). 0.0.0.0 is not allowed."
    done
}

# server_already_running <port>: true if a vLLM server process for this
# port exists or something is already listening on the port. Fail-closed:
# if neither pgrep nor ss is available, the guard cannot be checked, so
# the script refuses to launch rather than risk a duplicate server.
server_already_running() {
    local port="$1"
    if ! command -v pgrep >/dev/null 2>&1 && ! command -v ss >/dev/null 2>&1; then
        echo "ERROR: cannot check for an already-running server (neither pgrep nor ss is available)."
        echo "Install procps and iproute2, or stop any running server manually, then re-run."
        exit 1
    fi
    if command -v pgrep >/dev/null 2>&1 && \
       pgrep -f "vllm.entrypoints.openai.api_server.*--port ${port}" >/dev/null 2>&1; then
        return 0
    fi
    if command -v ss >/dev/null 2>&1 && \
       ss -ltn 2>/dev/null | grep -q ":${port} "; then
        return 0
    fi
    return 1
}

echo "=== Step 1: Serving settings ==="
action="$(prompt "Action (serve/healthcheck)" "serve")"
serve_backend="$(prompt "Backend (ollama/vllm)" "ollama")"
model_name="$(prompt "Model name (example — replace with your value)" "llama3.1:8b")"
serve_port="$(prompt "Serve port" "11434")"
gpu_host="$(prompt "GPU node hostname or internal IP (example — replace with your value)" "")"

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
config_path="$(prompt "Generation config path (example — replace with your value)" "/opt/airag/config.yaml")"
venv_path="$(prompt "Venv path (vLLM only, blank for Ollama)" "")"
serve_bind_ip=""
if [ "${serve_backend}" = "vllm" ]; then
    serve_bind_ip="$(prompt_bind_ip "vLLM bind address (internal interface IP, or 127.0.0.1 for localhost-only; never 0.0.0.0)")"
fi

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
chmod 640 "${config_path}"
echo "Config written to ${config_path}."

echo "=== Step 3: Serve ==="
if [ "${serve_backend}" = "ollama" ]; then
    sudo systemctl enable --now ollama
    curl -fsS "http://localhost:${serve_port}/api/tags" | head -c 300; echo
    echo "Ollama serving. Test: ollama run ${model_name} 'Say OK.'"
else
    if [ -z "${venv_path}" ]; then echo "ERROR: venv path required for vLLM."; exit 1; fi
    # Guard BEFORE any nohup launch: never start a second vLLM server
    # on a port that is already in use.
    if server_already_running "${serve_port}"; then
        echo "vLLM (or another server) is already running on port ${serve_port}."
        echo "Refusing to start a second instance. Use action healthcheck to probe it,"
        echo "or stop the existing server first, then re-run serve."
        exit 0
    fi
    nohup "${venv_path}/bin/python" -m vllm.entrypoints.openai.api_server \
        --host "${serve_bind_ip}" \
        --model "${model_name}" --port "${serve_port}" \
        --max-model-len "${context_length}" \
        > /var/log/airag/vllm.log 2>&1 &
    chmod 640 /var/log/airag/vllm.log 2>/dev/null || true
    sleep 15
    curl -fsS "http://${serve_bind_ip}:${serve_port}/v1/models" | head -c 300; echo
    echo "vLLM serving on ${serve_bind_ip}:${serve_port} (internal interface only; log: /var/log/airag/vllm.log)."
    echo "Reminder: this endpoint has no authentication of its own — keep it inside the internal"
    echo "network and front it with an authenticated TLS proxy before any non-local client uses it."
fi
echo "Done."

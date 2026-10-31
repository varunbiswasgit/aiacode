#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install vLLM (alternative to Ollama) into the project venv.
# Tracker task: P2-10
# Prompts:      venv_path, vllm_version (blank = latest)
# Prerequisites: Python venv created (02-install-python.sh), GPU node with
#               working NVIDIA driver. Run ON the GPU node.
#               NOTE: pick ONE production server - Ollama or vLLM, not both.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-vllm.log"
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
prompt_optional() {
    local value
    read -rp "$1 [$2]: " value
    echo "${value:-$2}"
}

echo "=== Step 1: Values ==="
venv_path="$(prompt "Virtual environment path" "/opt/airag/venv")"
vllm_version="$(prompt_optional "vLLM version (blank = latest)" "")"

echo "=== Step 2: Install ==="
if [ ! -x "${venv_path}/bin/python" ]; then
    echo "ERROR: venv not found at ${venv_path}. Run 02-install-python.sh first."; exit 1
fi
if [ -n "${vllm_version}" ]; then
    "${venv_path}/bin/pip" install "vllm==${vllm_version}"
else
    "${venv_path}/bin/pip" install vllm
fi

echo "=== Step 3: Verify ==="
"${venv_path}/bin/python" -c 'import vllm; print("vLLM", vllm.__version__)'
nvidia-smi -L
echo "Done. vLLM installed. Serve a model with 06-serve-language-model.sh."

#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Ollama (local LLM server) on the GPU node.
# Tracker task: P2-09
# Prompts:      ollama_port, ollama_models_dir
# Prerequisites: GPU node with working NVIDIA driver (01 script), sudo.
#               Run ON the GPU node. Model download is 06-install-language-model.sh.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-ollama.log"
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

echo "=== Step 1: Values ==="
ollama_port="$(prompt "Ollama API port" "11434")"
ollama_models_dir="$(prompt "Model storage directory" "/usr/share/ollama/.ollama/models")"

echo "=== Step 2: Install (idempotent) ==="
if command -v ollama &>/dev/null; then
    echo "Ollama already installed: $(ollama --version)"
else
    curl -fsSL https://ollama.com/install.sh | sh
fi

echo "=== Step 3: Configure and start ==="
sudo mkdir -p "$(dirname "${ollama_models_dir}")"
sudo systemctl set-environment OLLAMA_HOST="0.0.0.0:${ollama_port}" 2>/dev/null || true
sudo systemctl enable --now ollama
sleep 2

echo "=== Step 4: Verify ==="
ollama --version
curl -fsS "http://localhost:${ollama_port}/api/tags"
echo
echo "Done. Ollama listening on port ${ollama_port}. Next: 06-install-language-model.sh."

#!/usr/bin/env bash
# =============================================================================
# Purpose:      Download and install a language model on the GPU node.
# Tracker tasks: P6-02 (download model weights), P6-03 (quantization choice)
# Prompts:      model_source (ollama/huggingface), model_name,
#               huggingface_token (secret, only for gated HF models),
#               quantization (none/AWQ/GPTQ), ollama_port, models_dir
# Prerequisites: Ollama installed (02-install-ollama.sh) for ollama source,
#               or HF CLI (pip install huggingface_hub) for huggingface source.
#               Run ON the GPU node. The token is never echoed or saved.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-install-model.log"
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

echo "=== Step 1: Model source ==="
model_source="$(prompt "Model source (ollama/huggingface)" "ollama")"
model_name=""
huggingface_token=""
quantization="none"

case "${model_source}" in
    ollama)
        model_name="$(prompt "Ollama model name (e.g. llama3.1:8b, qwen2.5:14b)" "llama3.1:8b")"
        ollama_port="$(prompt "Ollama API port" "11434")"
        ;;
    huggingface)
        model_name="$(prompt "Hugging Face model id (e.g. meta-llama/Llama-3.1-8B-Instruct)" "meta-llama/Llama-3.1-8B-Instruct")"
        read -rsp "Hugging Face token (not echoed; blank if the model is public): " huggingface_token
        echo
        quantization="$(prompt "Quantization (none/AWQ/GPTQ)" "none")"
        models_dir="$(prompt "Local models directory" "/opt/airag/models")"
        ;;
    *)
        echo "ERROR: model source must be ollama or huggingface."; exit 1;;
esac

echo "=== Step 2: Download ==="
if [ "${model_source}" = "ollama" ]; then
    if ollama list 2>/dev/null | grep -q "${model_name}"; then
        echo "Model ${model_name} already present in Ollama."
    else
        ollama pull "${model_name}"
    fi
    ollama list
    echo "Smoke test:"
    curl -fsS "http://localhost:${ollama_port}/api/generate" -d "{\"model\": \"${model_name}\", \"prompt\": \"Say OK.\", \"stream\": false}" | head -c 300
    echo
else
    mkdir -p "${models_dir}"
    pip install -q huggingface_hub 2>/dev/null || python3 -m pip install -q huggingface_hub
    if [ -n "${huggingface_token}" ]; then
        export HF_TOKEN="${huggingface_token}"
    fi
    echo "Downloading ${model_name} (quantization: ${quantization})..."
    if [ "${quantization}" = "AWQ" ] || [ "${quantization}" = "GPTQ" ]; then
        echo "NOTE: for ${quantization} use the pre-quantized repo id (e.g. TheBloke/<model>-AWQ)."
        echo "Re-run and enter that repo id as the model name."
    fi
    python3 -c "from huggingface_hub import snapshot_download; print(snapshot_download('${model_name}', local_dir='${models_dir}/$(basename "${model_name}")'))"
    unset HF_TOKEN huggingface_token
    echo "Downloaded to ${models_dir}."
    echo "Verify size and checksum against the model card before serving."
fi
echo "Done. Next: 06-serve-language-model.sh."

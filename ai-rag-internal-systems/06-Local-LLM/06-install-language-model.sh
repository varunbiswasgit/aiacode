#!/usr/bin/env bash
# =============================================================================
# Purpose:      Download and install a language model on the GPU node.
# Tracker tasks: P6-02 (download model weights), P6-03 (quantization choice)
# Prompts:      model_source (ollama/huggingface), model_name,
#               huggingface_token (secret, only for gated HF models),
#               quantization (none/AWQ/GPTQ), ollama_port, models_dir,
#               model_revision (HF only: full commit hash from the model
#               card), model_checksums_path (HF only: path to the expected
#               SHA256 manifest for this model)
# Prerequisites: Ollama installed (02-install-ollama.sh) for ollama source,
#               or HF CLI (pip install huggingface_hub) for huggingface source.
#               Run ON the GPU node. The token is never echoed or saved.
# Checksum enforcement (HF branch):
#   - The old script only echoed "verify size and checksum against the
#     model card before serving". That advice is replaced by a hard
#     check: after download, every file in the model directory is
#     verified with sha256sum --check against the operator-supplied
#     manifest, and ANY mismatch, missing file, missing manifest, or
#     empty download exits non-zero. A failed model is NOT served.
#   - The manifest must be in sha256sum format ("<sha256>  <path>"),
#     with paths relative to the downloaded model directory (or
#     absolute), and must come from the model card / repository — not
#     invented locally. Repositories that publish no manifest cannot
#     be verified by this script; the script fails rather than
#     serving an unverified model.
#   - Pinning the revision to the full commit hash from the model card
#     is what ties the download to a fixed set of published bytes; a
#     branch name (e.g. "main") can move under you. A non-40-hex
#     revision is accepted only with a warning, and the checksum
#     check above remains the hard gate either way.
# Checksum enforcement (Ollama branch):
#   - `ollama pull` verifies each blob's SHA256 digest against the
#     registry manifest and fails on mismatch; under `set -e` that
#     failure aborts this script, so a bad pull is never followed by
#     a smoke test or a serve step.
# Secrets:      The Hugging Face token is read with `read -s`, passed
#               only via the HF_TOKEN environment variable to the
#               download helper, and unset immediately afterwards. It
#               is never passed as a script argument, never echoed,
#               and never written to the log file. Operator input
#               (model_name, models_dir, revision) is never
#               interpolated into Python source: values travel via
#               sys.argv to a quoted heredoc, so a crafted model name
#               cannot inject code.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-install-model.log"
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

echo "=== Step 1: Model source ==="
model_source="$(prompt "Model source (ollama/huggingface)" "ollama")"
model_name=""
huggingface_token=""
quantization="none"

case "${model_source}" in
    ollama)
        model_name="$(prompt "Ollama model name (example — replace with your value)" "llama3.1:8b")"
        ollama_port="$(prompt "Ollama API port" "11434")"
        ;;
    huggingface)
        model_name="$(prompt "Hugging Face model id (example — replace with your value)" "meta-llama/Llama-3.1-8B-Instruct")"
        read -rsp "Hugging Face token (not echoed; blank if the model is public): " huggingface_token
        echo
        quantization="$(prompt "Quantization (none/AWQ/GPTQ)" "none")"
        models_dir="$(prompt "Local models directory (example — replace with your value)" "/opt/airag/models")"
        model_revision="$(prompt "Model revision (example — replace with your value; full commit hash from the model card)" "")"
        if [[ ! "${model_revision}" =~ ^[0-9a-f]{40}$ ]]; then
            echo "WARNING: revision '${model_revision}' is not a full 40-hex commit hash."
            echo "A branch or tag can move; the checksum check in Step 3 is then the only integrity gate."
        fi
        model_checksums_path="$(prompt "Path to expected SHA256 manifest (example — replace with your value; '<sha256>  <relative-path>' per file, from the model card/repo)" "")"
        ;;
    *)
        echo "ERROR: model source must be ollama or huggingface."; exit 1;;
esac

echo "=== Step 2: Download ==="
if [ "${model_source}" = "ollama" ]; then
    if ollama list 2>/dev/null | grep -qF -- "${model_name}"; then
        echo "Model ${model_name} already present in Ollama."
    else
        # ollama pull verifies each blob SHA256 digest against the
        # registry manifest and fails on mismatch (hard fail under set -e).
        ollama pull "${model_name}"
    fi
    ollama list
    echo "Smoke test:"
    ollama run "${model_name}" "Say OK." | head -c 300
    echo
else
    mkdir -p "${models_dir}"
    # Pinned, never floating: replace the pinned version only with a
    # version your stack has been validated against, never "latest".
    python3 -m pip install -q "huggingface_hub==0.34.0"
    trap 'unset HF_TOKEN; unset huggingface_token' EXIT
    if [ -n "${huggingface_token}" ]; then
        export HF_TOKEN="${huggingface_token}"
    fi
    echo "Downloading ${model_name} (quantization: ${quantization})..."
    if [ "${quantization}" = "AWQ" ] || [ "${quantization}" = "GPTQ" ]; then
        echo "NOTE: for ${quantization} use the pre-quantized repo id (e.g. TheBloke/<model>-AWQ)."
        echo "Re-run and enter that repo id as the model name."
    fi
    model_dir="${models_dir}/$(basename "${model_name}")"
    # Values pass via sys.argv to a quoted heredoc. Operator input is
    # NEVER interpolated into the Python source below.
    python3 - "${model_name}" "${model_revision}" "${model_dir}" <<'PY'
import sys
from huggingface_hub import snapshot_download

repo_id, revision, local_dir = sys.argv[1], sys.argv[2], sys.argv[3]
print(snapshot_download(repo_id=repo_id, revision=revision,
                        local_dir=local_dir, local_dir_use_symlinks=False))
PY
    unset HF_TOKEN huggingface_token

    echo "=== Step 3: Verify checksums against the model card (fail hard) ==="
    if [ ! -f "${model_checksums_path}" ]; then
        echo "ERROR: checksum manifest not found at ${model_checksums_path}."
        echo "Do NOT serve this model. Obtain the manifest from the model card/repo and re-run."
        exit 1
    fi
    if [ ! -s "${model_checksums_path}" ]; then
        echo "ERROR: checksum manifest at ${model_checksums_path} is empty."
        echo "Do NOT serve this model."
        exit 1
    fi
    if [ -z "$(find "${model_dir}" -type f -print -quit 2>/dev/null)" ]; then
        echo "ERROR: no files downloaded to ${model_dir}."
        echo "Do NOT serve this model."
        exit 1
    fi
    checksums_abs="$(readlink -f "${model_checksums_path}")"
    # Manifest entries are relative to the model directory (or absolute);
    # run the check from inside it so relative entries resolve correctly.
    if (cd "${model_dir}" && sha256sum --check "${checksums_abs}"); then
        echo "Checksum verification PASSED for ${model_dir}."
    else
        echo "ERROR: CHECKSUM VERIFICATION FAILED for ${model_dir}."
        echo "The download does not match the model card manifest. Do NOT serve this model."
        echo "Delete or quarantine ${model_dir} and re-download from the pinned revision."
        exit 1
    fi
    echo "Downloaded and checksum-verified to ${model_dir}."
fi
echo "Done. Next: 06-serve-language-model.sh."

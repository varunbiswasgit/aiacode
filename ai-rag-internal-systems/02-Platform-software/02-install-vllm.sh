#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install vLLM (alternative to Ollama) into the project venv.
# Tracker task: P2-10
# Prompts:      venv_path, vllm_version (pinned, exact version only —
#               never blank, never "latest", never a range),
#               vllm_lockfile_path (optional: path to an existing
#               hash-verified requirements lockfile)
# Prerequisites: Python venv created (02-install-python.sh), GPU node with
#               working NVIDIA driver. Run ON the GPU node.
#               NOTE: pick ONE production server - Ollama or vLLM, not both.
# Version pinning:
#   - There is no unpinned install in this script. The default path
#     installs exactly "vllm==<pinned version>" and then writes a
#     pip freeze lockfile for reproducibility.
#   - A freeze file records versions; it does NOT hash-verify them.
#     For a hash-verified install, generate a lockfile with hashes
#     (e.g. pip-compile --generate-hashes) on a trusted build host,
#     supply its path at the lockfile prompt, and this script installs
#     it with pip --require-hashes. A lockfile that is missing, or
#     whose vllm pin differs from the prompted version, aborts the
#     install instead of falling back to an unpinned install.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-vllm.log"
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

prompt_version() { # prompt_version <label> <default> — exact X.Y.Z only
    local value
    while true; do
        read -rp "$1 [$2]: " value
        value="${value:-$2}"
        if [[ "${value}" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
            echo "${value}"; return
        fi
        echo "Enter an exact version (e.g. 0.31.0). 'latest', ranges, and wildcards are not accepted."
    done
}

echo "=== Step 1: Values ==="
venv_path="$(prompt "Virtual environment path (example — replace with your value)" "/opt/airag/venv")"
vllm_version="$(prompt_version "vLLM version (example — replace with your value; exact version only)" "0.31.0")"
vllm_lockfile_path=""
read -rp "Existing hash-verified requirements lockfile to install from (blank = install the exact pin and write a freeze lockfile): " vllm_lockfile_path
freeze_lockfile="${venv_path}/vllm-requirements.lock"

echo "=== Step 2: Install ==="
if [ ! -x "${venv_path}/bin/python" ]; then
    echo "ERROR: venv not found at ${venv_path}. Run 02-install-python.sh first."; exit 1
fi
if [ -n "${vllm_lockfile_path}" ]; then
    if [ ! -f "${vllm_lockfile_path}" ]; then
        echo "ERROR: lockfile not found at ${vllm_lockfile_path}. Refusing to fall back to an unpinned install."
        exit 1
    fi
    echo "Installing from hash-verified lockfile: ${vllm_lockfile_path}"
    "${venv_path}/bin/pip" install --require-hashes -r "${vllm_lockfile_path}"
else
    "${venv_path}/bin/pip" install "vllm==${vllm_version}"
    echo "Writing reproducibility lockfile (pip freeze; versions only, not hashes): ${freeze_lockfile}"
    "${venv_path}/bin/pip" freeze > "${freeze_lockfile}"
    chmod 640 "${freeze_lockfile}"
fi

echo "=== Step 3: Verify ==="
installed_version="$("${venv_path}/bin/python" -c 'import vllm; print(vllm.__version__)')"
echo "Installed vLLM version: ${installed_version}"
if [ "${installed_version}" != "${vllm_version}" ]; then
    echo "ERROR: installed vLLM ${installed_version} does not match the requested pin ${vllm_version}. Do not serve from this venv."
    exit 1
fi
nvidia-smi -L
echo "Done. vLLM ${installed_version} installed. Serve a model with 06-serve-language-model.sh."

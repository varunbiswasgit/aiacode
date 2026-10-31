#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Python 3.12 and create the project virtual environment.
# Tracker task: P2-02
# Prompts:      python_version, venv_path, project_owner_account
# Prerequisites: Ubuntu 24.04, sudo. Run ON the app node and the GPU node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-python.log"
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
python_version="$(prompt "Python version" "3.12")"
venv_path="$(prompt "Virtual environment path" "/opt/airag/venv")"
project_owner_account="$(prompt "Project owner account" "svc-airag-app")"

echo "=== Step 2: Install Python ==="
sudo apt update
sudo apt install -y "python${python_version}" "python${python_version}-venv" python3-pip

echo "=== Step 3: Create the venv (idempotent) ==="
sudo mkdir -p "$(dirname "${venv_path}")"
if [ -x "${venv_path}/bin/python" ]; then
    echo "Venv already exists at ${venv_path}."
else
    sudo "python${python_version}" -m venv "${venv_path}"
fi
sudo chown -R "${project_owner_account}:${project_owner_account}" "$(dirname "${venv_path}")"

echo "=== Step 4: Verify ==="
"${venv_path}/bin/python" --version
"${venv_path}/bin/pip" --version
echo "Done. Activate with: source ${venv_path}/bin/activate"

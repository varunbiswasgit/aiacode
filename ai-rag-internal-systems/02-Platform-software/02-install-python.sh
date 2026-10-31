#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Python 3.12 and create the project virtual environment.
# Tracker task: P2-02
# Prompts:      python_version, venv_path, venv_base, project_owner_account
# Prerequisites: Ubuntu 24.04, sudo. Run ON the app node and the GPU node.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-python.log"
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

echo "=== Step 1: Values ==="
python_version="$(prompt "Python version (example — replace with your value)" "3.12")"
venv_path="$(prompt "Virtual environment path (example — replace with your value)" "/opt/airag/venv")"
venv_base="$(prompt "Virtual environment base directory — the venv parent must be under this prefix (example — replace with your value)" "/opt/airag")"
project_owner_account="$(prompt "Project owner account (example — replace with your value)" "svc-airag-app")"

# Validate the venv parent path BEFORE any mkdir, venv creation, or
# recursive chown. A mistyped venv_path must never hand a system directory
# (e.g. /, /etc, /usr) to the service account via chown -R.
venv_base="${venv_base%/}"
venv_parent="$(dirname "${venv_path}")"
if [ -z "${venv_base}" ] || [ "${venv_base}" = "/" ]; then
    echo "ERROR: venv base '${venv_base}' is not a usable prefix (it must be a real directory below /). Refusing to continue."
    exit 1
fi
case "${venv_path}" in
    */../*|*/..|../*|..)
        echo "ERROR: venv path '${venv_path}' contains '..' components. Refusing to continue."
        exit 1
        ;;
esac
case "${venv_parent}" in
    "${venv_base}"|"${venv_base}"/*)
        # Parent is the base itself or a directory under it: allowed.
        ;;
    *)
        echo "ERROR: venv parent '${venv_parent}' is not under the prompted base '${venv_base}'. Refusing to chown anything outside that prefix."
        exit 1
        ;;
esac

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
# Safe by the validation above: dirname(venv_path) is venv_base or under it.
sudo chown -R "${project_owner_account}:${project_owner_account}" "$(dirname "${venv_path}")"

echo "=== Step 4: Verify ==="
"${venv_path}/bin/python" --version
"${venv_path}/bin/pip" --version
echo "Done. Activate with: source ${venv_path}/bin/activate"

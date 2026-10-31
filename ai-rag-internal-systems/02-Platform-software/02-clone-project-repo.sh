#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Git and clone the project repository onto the node.
# Tracker task: P2-03
# Prompts:      git_server_hostname, git_repository_path, project_dir,
#               project_owner_account, checkout_branch
# Prerequisites: Ubuntu 24.04, sudo, bare repo already created (00 script).
#               Run ON the app node.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-clone-repo.log"
sudo mkdir -p /var/log/airag
sudo touch "${LOG_FILE}" 2>/dev/null || true
sudo chmod 640 "${LOG_FILE}" 2>/dev/null || true
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

echo "=== Step 1: Repository details ==="
git_server_hostname="$(prompt "Git server hostname (example — replace with your value)" "git.internal")"
git_repository_path="$(prompt "Repository path on the server (example — replace with your value)" "/srv/git/airag.git")"
project_dir="$(prompt "Local project directory (example — replace with your value)" "/opt/airag")"
project_owner_account="$(prompt "Project owner account (example — replace with your value)" "svc-airag-app")"
checkout_branch="$(prompt "Branch to check out (example — replace with your value)" "main")"

echo "=== Step 2: Clone (idempotent) ==="
sudo apt install -y git
sudo mkdir -p "${project_dir}"
sudo chown "${project_owner_account}:${project_owner_account}" "${project_dir}"
if [ -d "${project_dir}/.git" ]; then
    echo "Already a clone; pulling latest."
    sudo -u "${project_owner_account}" git -C "${project_dir}" pull --ff-only || true
else
    sudo -u "${project_owner_account}" git clone "ssh://${git_server_hostname}${git_repository_path}" "${project_dir}"
fi
sudo -u "${project_owner_account}" git -C "${project_dir}" checkout "${checkout_branch}" 2>/dev/null || \
    echo "WARNING: branch ${checkout_branch} not found; staying on current branch."
sudo -u "${project_owner_account}" git -C "${project_dir}" status --short --branch
echo "Done. Project at ${project_dir}."

#!/usr/bin/env bash
# =============================================================================
# Purpose:      Create the AI-RAG project Git repository on the internal Git
#               server and clone it onto the admin workstation.
# Tracker task: P0-12
# Prompts:      git_server_hostname, git_repository_path, local_clone_path,
#               default_branch
# Prerequisites: SSH access to the internal Git server, git installed locally.
#               Run the server part on the Git server, the clone part anywhere.
# =============================================================================
set -euo pipefail

LOG_FILE="${HOME}/airag-create-repo.log"
exec > >(tee -a "${LOG_FILE}") 2>&1

prompt() { # prompt <label> <default>
    local value
    while true; do
        read -rp "$1 [$2]: " value
        value="${value:-$2}"
        if [ -n "${value}" ]; then echo "${value}"; return; fi
        echo "A value is required. Please try again."
    done
}

echo "=== Step 1: Where the repository will live ==="
git_server_hostname="$(prompt "Git server hostname" "git.internal")"
git_repository_path="$(prompt "Repository path on the server" "/srv/git/airag.git")"
local_clone_path="$(prompt "Local clone path" "${HOME}/airag")"
default_branch="$(prompt "Default branch name" "main")"

echo "=== Step 2: Create the bare repository (on the Git server) ==="
echo "Run this on ${git_server_hostname}:"
echo "  sudo mkdir -p ${git_repository_path}"
echo "  sudo git init --bare ${git_repository_path}"
read -rp "Press Enter once the bare repository exists on the server... "

echo "=== Step 3: Clone it locally ==="
if [ -d "${local_clone_path}/.git" ]; then
    echo "Already cloned at ${local_clone_path}; pulling latest."
    git -C "${local_clone_path}" pull --ff-only || true
else
    git clone "ssh://${git_server_hostname}${git_repository_path}" "${local_clone_path}"
fi

echo "=== Step 4: Seed the branching structure ==="
cd "${local_clone_path}"
git checkout -b "${default_branch}" 2>/dev/null || git checkout "${default_branch}"
for branch_name in develop docs; do
    git branch "${branch_name}" 2>/dev/null || echo "Branch ${branch_name} already exists."
done
mkdir -p docs prompts sql eval
touch docs/.gitkeep prompts/.gitkeep sql/.gitkeep eval/.gitkeep
git add -A
git commit -m "Seed repository structure" 2>/dev/null || echo "Nothing new to commit."
git push -u origin "${default_branch}" 2>/dev/null || echo "Push skipped (no remote access from here); push manually when ready."

echo "Done. Repository ready at ${local_clone_path} (branches: ${default_branch}, develop, docs)."

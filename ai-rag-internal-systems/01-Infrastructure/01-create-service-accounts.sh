#!/usr/bin/env bash
# =============================================================================
# Purpose:      Create the AI-RAG service accounts with least privilege.
# Tracker task: P1-10
# Prompts:      app_service_account, ingest_service_account, db_service_account,
#               add_to_docker_group (yes/no)
# Prerequisites: Ubuntu 24.04, sudo. Run ON each node (accounts are local;
#               create the same names on every node for consistency).
#               Passwords: set them in your vault afterwards - this script
#               locks password login and relies on SSH keys / AD.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-service-accounts.log"
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

echo "=== Step 1: Account names ==="
app_service_account="$(prompt "App service account" "svc-airag-app")"
ingest_service_account="$(prompt "Ingest service account" "svc-airag-ingest")"
db_service_account="$(prompt "Database service account" "svc-airag-db")"
add_to_docker_group="$(prompt "Add the app account to the docker group? (yes/no)" "yes")"

echo "=== Step 2: Create accounts (idempotent) ==="
for account_name in "${app_service_account}" "${ingest_service_account}" "${db_service_account}"; do
    if id "${account_name}" &>/dev/null; then
        echo "Account ${account_name} already exists - skipping."
    else
        sudo useradd -m -s /bin/bash "${account_name}"
        sudo passwd -l "${account_name}"   # lock password login; use SSH keys
        echo "Created ${account_name}."
    fi
done

if [ "${add_to_docker_group}" = "yes" ] && getent group docker &>/dev/null; then
    sudo usermod -aG docker "${app_service_account}"
    echo "Added ${app_service_account} to docker group."
fi

echo "=== Step 3: Verify ==="
id "${app_service_account}"; id "${ingest_service_account}"; id "${db_service_account}"
echo "Done. Store any passwords/keys for these accounts in your vault (not in this log)."

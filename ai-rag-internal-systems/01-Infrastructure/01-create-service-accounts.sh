#!/usr/bin/env bash
# =============================================================================
# Purpose:      Create the AI-RAG service accounts with least privilege.
#               No account is given Docker access: docker group membership
#               is root-equivalent, so containers run from root-owned
#               systemd units instead (see 08-create-systemd-services.sh).
#               If the app account is already in the docker group (e.g.
#               from an earlier run), this script removes it.
# Tracker task: P1-10
# Prompts:      app_service_account, ingest_service_account,
#               db_service_account
# Prerequisites: Ubuntu 24.04, sudo. Run ON each node (accounts are local;
#               create the same names on every node for consistency).
#               Passwords: set them in your vault afterwards - this script
#               locks password login and relies on SSH keys / AD.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-service-accounts.log"
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

echo "=== Step 1: Account names ==="
app_service_account="$(prompt "App service account" "svc-airag-app")"
ingest_service_account="$(prompt "Ingest service account" "svc-airag-ingest")"
db_service_account="$(prompt "Database service account" "svc-airag-db")"

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

echo "=== Step 3: Ensure the app account has NO Docker access ==="
if getent group docker &>/dev/null && id -nG "${app_service_account}" | tr ' ' '\n' | grep -qx docker; then
    sudo gpasswd -d "${app_service_account}" docker
    echo "Removed ${app_service_account} from the docker group."
else
    echo "${app_service_account} is not in the docker group (by design)."
fi

echo "=== Step 4: Verify ==="
id "${app_service_account}"; id "${ingest_service_account}"; id "${db_service_account}"
if getent group docker &>/dev/null && id -nG "${app_service_account}" | tr ' ' '\n' | grep -qx docker; then
    echo "ERROR: ${app_service_account} is still in the docker group." >&2
    exit 1
fi
echo "Done. Store any passwords/keys for these accounts in your vault (not in this log)."
echo "Containers run from root-owned systemd units; the app account has no Docker access."

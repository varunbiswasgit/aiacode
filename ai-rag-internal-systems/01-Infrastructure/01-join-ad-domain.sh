#!/usr/bin/env bash
# =============================================================================
# Purpose:      Join the node to Active Directory and verify domain login.
# Tracker task: P1-09
# Prompts:      ad_domain_name, ad_admin_username, ad_admin_password (secret),
#               test_domain_username
# Prerequisites: Ubuntu 24.04, sudo, network route to a domain controller.
#               Run ON each node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-join-ad-domain.log"
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

echo "=== Step 1: Domain values ==="
ad_domain_name="$(prompt "Active Directory domain" "corp.example.com")"
ad_admin_username="$(prompt "Domain admin username" "administrator")"
read -rsp "Domain admin password (not echoed, never logged): " ad_admin_password
echo

echo "=== Step 2: Install AD client packages ==="
sudo apt update
sudo apt install -y realmd sssd-ad adcli sssd-tools packagekit

echo "=== Step 3: Join the domain ==="
if realm list | grep -qi "${ad_domain_name}"; then
    echo "Already joined to ${ad_domain_name}."
else
    echo "${ad_admin_password}" | sudo realm join --user="${ad_admin_username}" "${ad_domain_name}" || {
        echo "ERROR: realm join failed. Check the domain, credentials, DNS and time sync."; exit 1; }
fi
unset ad_admin_password

echo "=== Step 4: Verify ==="
realm list
test_domain_username="$(prompt "Test domain username (without @domain)" "testuser")"
id "${test_domain_username}@${ad_domain_name}" && echo "Domain login lookup OK."
echo "Done. Node joined to ${ad_domain_name}."

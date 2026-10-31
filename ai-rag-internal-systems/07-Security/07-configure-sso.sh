#!/usr/bin/env bash
# =============================================================================
# Purpose:      Write the app's SSO configuration (OIDC or LDAP) from prompts.
# Tracker task: P7-01 (wire AD/LDAP SSO)
# Prompts:      sso_type (oidc/ldap), provider_url, config_path,
#               client_id, client_secret (secret, OIDC only),
#               ldap_bind_dn, ldap_bind_password (secret, LDAP only),
#               test_username
# Prerequisites: Ubuntu 24.04, app code that reads this config file.
#               Run ON the app node. Secrets are never echoed or logged.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-sso.log"
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

echo "=== Step 1: SSO type ==="
sso_type="$(prompt "SSO type (oidc/ldap) (example — replace with your value)" "oidc")"
provider_url="$(prompt "Provider URL (OIDC issuer or LDAP server URL) (example — replace with your value)" "https://login.internal")"
config_path="$(prompt "Config file path (example — replace with your value)" "/opt/airag/sso-config.yaml")"

echo "=== Step 2: Credentials (never logged) ==="
if [ "${sso_type}" = "oidc" ]; then
    client_id="$(prompt "OIDC client ID (example — replace with your value)" "airag-app")"
    read -rsp "OIDC client secret (not echoed): " client_secret
    echo
    cat > "${config_path}" <<EOF
# AI-RAG SSO configuration (P7-01) - secrets live here; restrict permissions!
sso_type: oidc
issuer_url: ${provider_url}
client_id: ${client_id}
client_secret: ${client_secret}
EOF
    unset client_secret
elif [ "${sso_type}" = "ldap" ]; then
    ldap_bind_dn="$(prompt "LDAP bind DN (example — replace with your value)" "CN=svc-airag-app,OU=Service,DC=internal")"
    read -rsp "LDAP bind password (not echoed): " ldap_bind_password
    echo
    cat > "${config_path}" <<EOF
# AI-RAG SSO configuration (P7-01) - secrets live here; restrict permissions!
sso_type: ldap
ldap_url: ${provider_url}
bind_dn: ${ldap_bind_dn}
bind_password: ${ldap_bind_password}
EOF
    unset ldap_bind_password
else
    echo "ERROR: sso_type must be oidc or ldap."; exit 1
fi
chmod 600 "${config_path}"
echo "Config written to ${config_path} (mode 600)."

echo "=== Step 3: Verify ==="
test_username="$(prompt "Test username for login check (example — replace with your value)" "testuser")"
echo "Log in to the app as ${test_username} and confirm SSO succeeds for 2 test users."
echo "Done. If login fails, check the issuer URL / bind DN and the app log."

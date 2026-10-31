#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Nginx and set up HTTPS with an internal-CA certificate.
# Tracker task: P2-11
# Prompts:      app_domain_name, app_upstream_port, cert_country, cert_org
#               (cert_common_name is derived from app_domain_name, not
#               prompted; there is no ca_submission_method prompt — the CSR
#               is submitted to the internal CA manually in Step 3.)
# Prerequisites: Ubuntu 24.04, sudo, app port known. Run ON the app node.
#               The CSR must be signed by the internal CA team; this script
#               pauses for that step and resumes when the cert is in place.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-nginx-tls.log"
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
app_domain_name="$(prompt "Application domain name (example — replace with your value)" "airag.internal")"
app_upstream_port="$(prompt "Application upstream port (FastAPI) (example — replace with your value)" "8000")"
cert_country="$(prompt "Certificate country code (example — replace with your value)" "US")"
cert_org="$(prompt "Certificate organisation (example — replace with your value)" "Internal")"
cert_common_name="${app_domain_name}"

echo "=== Step 2: Install Nginx ==="
sudo apt update
sudo apt install -y nginx openssl
sudo systemctl enable --now nginx

echo "=== Step 3: Create key and CSR ==="
sudo mkdir -p /etc/ssl/private /etc/ssl/certs
if [ ! -f /etc/ssl/private/airag.key ]; then
    sudo openssl req -new -newkey rsa:2048 -nodes \
        -keyout /etc/ssl/private/airag.key \
        -out /tmp/airag.csr \
        -subj "/C=${cert_country}/O=${cert_org}/CN=${cert_common_name}"
    echo "CSR created at /tmp/airag.csr - submit it to the internal CA now."
else
    echo "Key already exists; reusing it."
fi
read -rp "Press Enter once the signed certificate is saved at /etc/ssl/certs/airag.crt ... "
if [ ! -f /etc/ssl/certs/airag.crt ]; then
    echo "ERROR: /etc/ssl/certs/airag.crt not found. Place the cert and re-run."; exit 1
fi
# The CSR has served its purpose once the signed certificate is in place:
# delete it so the signing request (with the certificate subject details)
# is not left lying around on disk. This runs only after the cert check
# above passes, so a failed run keeps the CSR for resubmission.
if [ -f /tmp/airag.csr ]; then
    sudo rm -f /tmp/airag.csr
    echo "CSR deleted (/tmp/airag.csr)."
fi

echo "=== Step 4: Nginx site ==="
sudo tee /etc/nginx/sites-available/airag > /dev/null <<EOF
server {
    listen 443 ssl;
    server_name ${app_domain_name};
    ssl_certificate /etc/ssl/certs/airag.crt;
    ssl_certificate_key /etc/ssl/private/airag.key;
    # TLS 1.2 and 1.3 only; older protocols (SSLv3, TLS 1.0/1.1) are disabled.
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_prefer_server_ciphers on;
    ssl_ciphers ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384:ECDHE-ECDSA-CHACHA20-POLY1305:ECDHE-RSA-CHACHA20-POLY1305;
    add_header Strict-Transport-Security "max-age=63072000; includeSubDomains" always;
    location / {
        proxy_pass http://127.0.0.1:${app_upstream_port};
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }
}
server {
    listen 80;
    server_name ${app_domain_name};
    return 301 https://\$host\$request_uri;
}
EOF
sudo ln -sf /etc/nginx/sites-available/airag /etc/nginx/sites-enabled/airag
sudo nginx -t
sudo systemctl reload nginx

echo "=== Step 5: Verify ==="
# Primary check: full TLS validation, no -k. --resolve maps the certificate's
# domain name to localhost so the hostname in the cert matches. For this to
# pass, the internal CA certificate must be installed in the system trust
# store (e.g. under /usr/local/share/ca-certificates, then
# update-ca-certificates) before re-running this check.
if curl -fsS "https://${app_domain_name}/" --resolve "${app_domain_name}:443:127.0.0.1" -o /dev/null -w "Local HTTPS status (certificate validated): %{http_code}\n"; then
    echo "TLS verification passed: certificate validated against the system trust store."
else
    echo "WARNING: full TLS validation did not pass. This can mean the internal CA certificate is not yet installed in the system trust store, or the app is not yet running on port ${app_upstream_port}."
    echo "PRE-CA FALLBACK ONLY: the -k check below does NOT validate the certificate; it only shows whether Nginx answers on HTTPS. Re-run the validated check above once the internal CA certificate is installed."
    curl -ks "https://localhost/" -o /dev/null -w "Local HTTPS status (UNVALIDATED, -k fallback): %{http_code}\n" || echo "App not yet running on port ${app_upstream_port} - HTTPS proxy itself is configured."
fi
echo "Done. Browse to https://${app_domain_name} once the app is running."

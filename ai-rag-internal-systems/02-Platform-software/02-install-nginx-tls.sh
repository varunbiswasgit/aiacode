#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Nginx and set up HTTPS with an internal-CA certificate.
# Tracker task: P2-11
# Prompts:      app_domain_name, app_upstream_port, cert_country, cert_org,
#               cert_common_name, ca_submission_method
# Prerequisites: Ubuntu 24.04, sudo, app port known. Run ON the app node.
#               The CSR must be signed by the internal CA team; this script
#               pauses for that step and resumes when the cert is in place.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-nginx-tls.log"
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
app_domain_name="$(prompt "Application domain name" "airag.internal")"
app_upstream_port="$(prompt "Application upstream port (FastAPI)" "8000")"
cert_country="$(prompt "Certificate country code" "US")"
cert_org="$(prompt "Certificate organisation" "Internal")"
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

echo "=== Step 4: Nginx site ==="
sudo tee /etc/nginx/sites-available/airag > /dev/null <<EOF
server {
    listen 443 ssl;
    server_name ${app_domain_name};
    ssl_certificate /etc/ssl/certs/airag.crt;
    ssl_certificate_key /etc/ssl/private/airag.key;
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
curl -ks "https://localhost/" -o /dev/null -w "Local HTTPS status: %{http_code}\n" || echo "App not yet running on port ${app_upstream_port} - HTTPS proxy itself is configured."
echo "Done. Browse to https://${app_domain_name} once the app is running."

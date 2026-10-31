#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install the Prometheus node_exporter monitoring agent.
# Tracker task: P1-17
# Prompts:      node_exporter_version, install_dir, metrics_port
# Prerequisites: Ubuntu 24.04, sudo, internet or internal mirror.
#               Run ON each node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-monitoring-agent.log"
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

echo "=== Step 1: Version and paths ==="
node_exporter_version="$(prompt "node_exporter version" "1.8.2")"
install_dir="$(prompt "Install directory" "/usr/local/bin")"
metrics_port="$(prompt "Metrics port" "9100")"
download_url="https://github.com/prometheus/node_exporter/releases/download/v${node_exporter_version}/node_exporter-${node_exporter_version}.linux-amd64.tar.gz"

echo "=== Step 2: Install (idempotent) ==="
if [ -x "${install_dir}/node_exporter" ]; then
    echo "node_exporter already installed; checking version."
    "${install_dir}/node_exporter" --version | head -1
else
    cd /tmp
    curl -fsSLO "${download_url}"
    tar xvfz "node_exporter-${node_exporter_version}.linux-amd64.tar.gz"
    sudo mv "node_exporter-${node_exporter_version}.linux-amd64/node_exporter" "${install_dir}/"
    rm -rf "/tmp/node_exporter-${node_exporter_version}.linux-amd64"*
fi

echo "=== Step 3: Service ==="
sudo tee /etc/systemd/system/node_exporter.service > /dev/null <<EOF
[Unit]
Description=Prometheus node_exporter
After=network.target

[Service]
User=nobody
ExecStart=${install_dir}/node_exporter --web.listen-address=:${metrics_port}
Restart=always

[Install]
WantedBy=multi-user.target
EOF
sudo systemctl daemon-reload
sudo systemctl enable --now node_exporter

echo "=== Step 4: Verify ==="
sleep 2
curl -fsS "http://localhost:${metrics_port}/metrics" | head -3
echo "Done. Metrics on port ${metrics_port}; point Prometheus at this node."

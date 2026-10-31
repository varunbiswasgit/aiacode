#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install the Prometheus node_exporter monitoring agent.
# Tracker task: P1-17
# Prompts:      node_exporter_version, install_dir, metrics_port,
#               expected_sha256 (optional)
# Prerequisites: Ubuntu 24.04, sudo, internet or internal mirror.
#               Run ON each node.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-monitoring-agent.log"
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

prompt_optional() { # prompt_optional <label> — empty answer is allowed
    local value
    read -rp "$1: " value
    echo "${value}"
}

echo "=== Step 1: Version and paths ==="
node_exporter_version="$(prompt "node_exporter version (example — replace with your value)" "1.8.2")"
install_dir="$(prompt "Install directory (example — replace with your value)" "/usr/local/bin")"
metrics_port="$(prompt "Metrics port (example — replace with your value)" "9100")"
expected_sha256="$(prompt_optional "Expected SHA256 checksum for this release, from the publisher's release page (optional — leave blank to verify against the release's published sha256sums.txt instead)")"

# Validate the version BEFORE it is used in the download URL or in any
# rm -rf glob below. Anything outside strict X.Y.Z (digits and dots only)
# is refused, so a mistyped or malicious value can never expand the glob
# into unrelated files.
if ! [[ "${node_exporter_version}" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "ERROR: node_exporter version '${node_exporter_version}' is not a plain X.Y.Z version. Refusing to continue."
    exit 1
fi
if ! [[ "${metrics_port}" =~ ^[0-9]+$ ]]; then
    echo "ERROR: metrics port '${metrics_port}' is not numeric. Refusing to continue."
    exit 1
fi
if [ -n "${expected_sha256}" ] && ! [[ "${expected_sha256}" =~ ^[0-9a-fA-F]{64}$ ]]; then
    echo "ERROR: expected SHA256 checksum is not 64 hex characters. Refusing to continue."
    exit 1
fi
expected_sha256="${expected_sha256,,}" # normalise case: sha256sum prints lowercase

archive="node_exporter-${node_exporter_version}.linux-amd64.tar.gz"
download_url="https://github.com/prometheus/node_exporter/releases/download/v${node_exporter_version}/${archive}"
checksum_url="https://github.com/prometheus/node_exporter/releases/download/v${node_exporter_version}/sha256sums.txt"

echo "=== Step 2: Install (idempotent) ==="
if [ -x "${install_dir}/node_exporter" ]; then
    echo "node_exporter already installed; checking version."
    "${install_dir}/node_exporter" --version | head -1
else
    cd /tmp
    curl -fsSLO "${download_url}"
    # Verify the release checksum BEFORE extracting or installing anything.
    # Fail (exit 1) on any mismatch; the archive is never installed unverified.
    # If the operator supplied an expected checksum, it is checked directly
    # (independent of the downloaded checksum file). Otherwise the release's
    # published sha256sums.txt is used; --ignore-missing checks only the
    # archive actually downloaded.
    if [ -n "${expected_sha256}" ]; then
        actual_sha256="$(sha256sum "${archive}" | awk '{print $1}')"
        if [ "${actual_sha256}" != "${expected_sha256}" ]; then
            echo "ERROR: SHA256 mismatch for ${archive}; refusing to install."
            echo "Expected: ${expected_sha256}"
            echo "Actual:   ${actual_sha256}"
            rm -f "${archive}"
            exit 1
        fi
        echo "SHA256 verified against the operator-supplied expected checksum."
    else
        curl -fsSLO "${checksum_url}"
        if ! sha256sum --check --ignore-missing sha256sums.txt; then
            echo "ERROR: SHA256 verification failed for ${archive}; refusing to install."
            rm -f "${archive}" sha256sums.txt
            exit 1
        fi
        echo "SHA256 verified against the release's published sha256sums.txt."
    fi
    tar xvfz "${archive}"
    sudo mv "node_exporter-${node_exporter_version}.linux-amd64/node_exporter" "${install_dir}/"
    rm -rf "/tmp/node_exporter-${node_exporter_version}.linux-amd64"* "${archive}" sha256sums.txt
fi

echo "=== Step 3: Service ==="
sudo tee /etc/systemd/system/node_exporter.service > /dev/null <<EOF
[Unit]
Description=Prometheus node_exporter
After=network.target

[Service]
User=nobody
# Metrics are bound to localhost only: the endpoint is not exposed to the
# network. Scrape it from a collector running on this node (or via an SSH
# tunnel / local agent); do not open the port to remote scrapers.
ExecStart=${install_dir}/node_exporter --web.listen-address=127.0.0.1:${metrics_port}
Restart=always

[Install]
WantedBy=multi-user.target
EOF
sudo systemctl daemon-reload
sudo systemctl enable --now node_exporter

echo "=== Step 4: Verify ==="
sleep 2
curl -fsS "http://localhost:${metrics_port}/metrics" | head -3
echo "Done. Metrics are bound to localhost on port ${metrics_port}; scrape them from this node only."

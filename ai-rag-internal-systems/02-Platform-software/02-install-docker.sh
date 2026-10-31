#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Docker Engine and Compose, with optional internal
#               registry mirror support. No service account is given Docker
#               access: docker group membership is root-equivalent, so
#               containers are started from root-owned systemd units
#               instead (see 08-create-systemd-services.sh).
# Tracker tasks: P2-01 (Docker), P2-12 (internal registry mirror)
# Prompts:      install_method (package/installer),
#               installer_sha256_checksum (installer method only),
#               use_internal_registry (yes/no), internal_registry_url
#               installer_sha256_checksum (example — replace with your
#               value): the published SHA-256 of the installer script for
#               the release being installed, e.g. sha256:<64 hex digits>,
#               obtained from Docker's published documentation or your
#               internal mirror before running this script. If you do not
#               have a published checksum to verify against, choose the
#               package method instead — the installer is never run
#               unverified.
#               internal_registry_url (example — replace with your value):
#               your internal registry hostname, e.g. registry.internal
#               is only a placeholder format — enter your real value.
# Prerequisites: Ubuntu 24.04, sudo, internet or internal mirror.
#               Run ON each node.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-docker.log"
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

prompt_required() { # prompt_required <label>
    local value
    while true; do
        read -rp "$1: " value
        if [ -n "${value}" ]; then echo "${value}"; return; fi
        echo "A value is required. Please try again."
    done
}

prompt_checksum() { # prompt_checksum <label>
    local value
    while true; do
        read -rp "$1: " value
        value="${value#sha256:}"
        if [[ "${value}" =~ ^[0-9a-fA-F]{64}$ ]]; then
            echo "${value}"; return
        fi
        echo "Enter a SHA-256 checksum: 64 hex digits, optionally prefixed with 'sha256:'."
    done
}

echo "=== Step 1: Options ==="
install_method=""
while [ "${install_method}" != "package" ] && [ "${install_method}" != "installer" ]; do
    install_method="$(prompt "Install method: distro package or Docker installer script? (package/installer)" "package")"
done
installer_sha256_checksum=""
if [ "${install_method}" = "installer" ]; then
    installer_sha256_checksum="$(prompt_checksum "Published SHA-256 of the installer script (example — replace with your value; sha256:<64 hex digits>)")"
fi
use_internal_registry="$(prompt "Use an internal container registry mirror? (yes/no)" "no")"
internal_registry_url=""
if [ "${use_internal_registry}" = "yes" ]; then
    internal_registry_url="$(prompt_required "Internal registry URL (example — replace with your value)")"
fi

echo "=== Step 2: Install Docker ==="
if command -v docker &>/dev/null; then
    echo "Docker already installed: $(docker --version)"
elif [ "${install_method}" = "package" ]; then
    echo "Installing Docker from the distro packages."
    sudo apt-get update
    sudo apt-get install -y docker.io docker-compose-v2
else
    echo "Downloading the Docker installer script for verification (it is NOT piped to a shell)."
    installer_file="$(mktemp /tmp/get-docker.XXXXXX.sh)"
    chmod 600 "${installer_file}"
    curl -fsSL -o "${installer_file}" https://get.docker.com
    if ! printf '%s  %s\n' "${installer_sha256_checksum}" "${installer_file}" | sha256sum -c -; then
        echo "ERROR: installer checksum mismatch. The downloaded script will NOT be run." >&2
        rm -f "${installer_file}"
        exit 1
    fi
    echo "Installer checksum verified. Running the LOCAL copy."
    sudo sh "${installer_file}"
    rm -f "${installer_file}"
fi
sudo systemctl enable --now docker
echo "No service account is added to the docker group (docker group membership is root-equivalent)."

echo "=== Step 3: Internal registry mirror (optional) ==="
if [ "${use_internal_registry}" = "yes" ]; then
    sudo mkdir -p /etc/docker
    sudo tee /etc/docker/daemon.json > /dev/null <<EOF
{
  "registry-mirrors": ["https://${internal_registry_url}"]
}
EOF
    sudo systemctl restart docker
    echo "Registry mirror set to ${internal_registry_url}."
fi

echo "=== Step 4: Verify ==="
sudo docker run --rm hello-world
sudo docker compose version
echo "Done. Docker ready. Containers are started from root-owned systemd units; no service account has Docker access."

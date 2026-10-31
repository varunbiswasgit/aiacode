#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Docker Engine and Compose, with optional internal
#               registry mirror support.
# Tracker tasks: P2-01 (Docker), P2-12 (internal registry mirror)
# Prompts:      docker_user_account, use_internal_registry (yes/no),
#               internal_registry_url
# Prerequisites: Ubuntu 24.04, sudo, internet or internal mirror.
#               Run ON each node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-docker.log"
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

echo "=== Step 1: Options ==="
docker_user_account="$(prompt "Service account to add to the docker group" "svc-airag-app")"
use_internal_registry="$(prompt "Use an internal container registry mirror? (yes/no)" "no")"
internal_registry_url=""
if [ "${use_internal_registry}" = "yes" ]; then
    internal_registry_url="$(prompt "Internal registry URL" "registry.internal")"
fi

echo "=== Step 2: Install Docker ==="
if command -v docker &>/dev/null; then
    echo "Docker already installed: $(docker --version)"
else
    curl -fsSL https://get.docker.com | sh
fi
sudo systemctl enable --now docker
if id "${docker_user_account}" &>/dev/null; then
    sudo usermod -aG docker "${docker_user_account}"
fi

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
docker compose version
echo "Done. Docker ready; log out/in for group membership to take effect."

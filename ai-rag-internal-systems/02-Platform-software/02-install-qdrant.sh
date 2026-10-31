#!/usr/bin/env bash
# =============================================================================
# Purpose:      Run Qdrant (vector store) as a Docker container with a
#               persistent volume.
# Tracker task: P2-07
# Prompts:      qdrant_http_port, qdrant_grpc_port, qdrant_volume_name,
#               qdrant_image
# Prerequisites: Docker installed (02-install-docker.sh). Run ON the DB node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-qdrant.log"
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
qdrant_http_port="$(prompt "Qdrant HTTP port" "6333")"
qdrant_grpc_port="$(prompt "Qdrant gRPC port" "6334")"
qdrant_volume_name="$(prompt "Docker volume name" "qdrant_storage")"
qdrant_image="$(prompt "Qdrant image" "qdrant/qdrant")"

echo "=== Step 2: Start Qdrant (idempotent) ==="
docker pull "${qdrant_image}"
if docker ps -a --format '{{.Names}}' | grep -qx qdrant; then
    docker start qdrant
    echo "Existing qdrant container started."
else
    docker run -d --name qdrant --restart unless-stopped \
        -p "${qdrant_http_port}:6333" -p "${qdrant_grpc_port}:6334" \
        -v "${qdrant_volume_name}:/qdrant/storage" \
        "${qdrant_image}"
fi

echo "=== Step 3: Verify ==="
sleep 3
curl -fsS "http://localhost:${qdrant_http_port}/" | head -c 300; echo
docker ps --filter name=qdrant --format '{{.Names}} {{.Status}}'
echo "Done. Qdrant on HTTP port ${qdrant_http_port}."

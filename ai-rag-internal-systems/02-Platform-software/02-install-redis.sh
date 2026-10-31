#!/usr/bin/env bash
# =============================================================================
# Purpose:      Run Redis (cache + job queue) as a Docker container with
#               append-only persistence.
# Tracker task: P2-08
# Prompts:      redis_port, redis_volume_name, redis_image
# Prerequisites: Docker installed (02-install-docker.sh). Run ON the DB node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-redis.log"
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
redis_port="$(prompt "Redis port" "6379")"
redis_volume_name="$(prompt "Docker volume name" "redis_data")"
redis_image="$(prompt "Redis image" "redis:7")"

echo "=== Step 2: Start Redis (idempotent) ==="
docker pull "${redis_image}"
if docker ps -a --format '{{.Names}}' | grep -qx redis; then
    docker start redis
    echo "Existing redis container started."
else
    docker run -d --name redis --restart unless-stopped \
        -p "${redis_port}:6379" -v "${redis_volume_name}:/data" \
        "${redis_image}" --appendonly yes
fi

echo "=== Step 3: Verify ==="
docker exec redis redis-cli ping
echo "Done. Redis on port ${redis_port} with persistence enabled."

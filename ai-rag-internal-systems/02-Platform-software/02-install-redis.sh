#!/usr/bin/env bash
# =============================================================================
# Purpose:      Run Redis (cache + job queue) as a Docker container with
#               append-only persistence, a required password, and its port
#               published only on the internal interface.
# Tracker task: P2-08
# Prompts:      redis_bind_ip, redis_port, redis_volume_name, redis_image,
#               redis_password (secret)
#               redis_bind_ip (example — replace with your value): the IP
#               of this node's internal interface, or 127.0.0.1 for a
#               localhost-only deployment. 0.0.0.0 is rejected.
#               redis_image (example — replace with your value): a
#               digest-pinned reference, e.g.
#               redis@sha256:<64 hex digits>. A mutable major tag such as
#               `redis:7` and `latest` are rejected; the operator supplies
#               the digest they verified for the release they intend to run.
# Secrets:      redis_password is read with `read -s`, written to a
#               mode-600 config file, and never passed as a script argument,
#               never echoed, and never written to the log file. The
#               container is started with a config-file path (not the
#               password), and verification uses the REDISCLI_AUTH
#               environment variable, not a command-line argument.
# Prerequisites: Docker installed (02-install-docker.sh). Run ON the DB node.
#               Containers are started by the operator/root, not by the app
#               service account (which has no Docker access).
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-redis.log"
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

prompt_bind_ip() { # prompt_bind_ip <label>
    local value
    while true; do
        read -rp "$1: " value
        if [[ "${value}" =~ ^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+$ ]] && [ "${value}" != "0.0.0.0" ]; then
            echo "${value}"; return
        fi
        echo "Enter an IPv4 address for the internal interface (or 127.0.0.1 for localhost-only). 0.0.0.0 is not allowed."
    done
}

prompt_image() { # prompt_image <label>
    local value
    while true; do
        read -rp "$1: " value
        if [[ "${value}" =~ ^[^[:space:]]+@sha256:[0-9a-fA-F]{64}$ ]]; then
            echo "${value}"; return
        fi
        echo "A digest-pinned image reference is required, e.g. redis@sha256:<64 hex digits>. Mutable tags such as 'redis:7' or 'latest' are not accepted."
    done
}

prompt_secret() { # prompt_secret <label> — reads twice, never echoes
    local first second
    while true; do
        read -rsp "$1: " first; echo
        read -rsp "$1 (confirm): " second; echo
        if [ -z "${first}" ]; then
            echo "A value is required. Please try again."
        elif [ "${first}" != "${second}" ]; then
            echo "Values do not match. Please try again."
        elif [[ "${first}" =~ [[:space:]] ]]; then
            echo "The password must not contain whitespace (it is stored in a Redis config file). Please try again."
        else
            printf '%s' "${first}"; return
        fi
    done
}

echo "=== Step 1: Values ==="
redis_bind_ip="$(prompt_bind_ip "Redis bind address (internal interface IP, or 127.0.0.1 for localhost-only; never 0.0.0.0)")"
redis_port="$(prompt "Redis port" "6379")"
redis_volume_name="$(prompt "Docker volume name" "redis_data")"
redis_image="$(prompt_image "Redis image (example — replace with your value; digest-pinned, e.g. redis@sha256:<64 hex digits>)")"
redis_password="$(prompt_secret "Redis password (requirepass)")"

echo "=== Step 2: Store the password in a mode-600 config file ==="
redis_config_file="/etc/airag/redis.conf"
sudo mkdir -p /etc/airag
sudo chmod 750 /etc/airag
{
    printf 'protected-mode yes\n'
    printf 'appendonly yes\n'
    printf 'requirepass %s\n' "${redis_password}"
} | sudo tee "${redis_config_file}" > /dev/null
sudo chmod 600 "${redis_config_file}"
echo "Redis config stored in ${redis_config_file} (mode 600)."
# Note: no `bind` line is set inside the container on purpose — the
# container must listen on its own interface for the published port to
# work. Exposure is controlled on the host by the -p bind address above,
# never by publishing on all interfaces.

echo "=== Step 3: Start Redis (idempotent; existing container is recreated) ==="
docker pull "${redis_image}"
if docker ps -a --format '{{.Names}}' | grep -qx redis; then
    # Recreate, rather than merely restarting, so a pre-existing container
    # without a password, a pinned digest, or an internal-only bind cannot
    # survive this run. The named volume (data) is preserved.
    docker rm -f redis
fi
docker run -d --name redis --restart unless-stopped \
    -p "${redis_bind_ip}:${redis_port}:6379" \
    -v "${redis_volume_name}:/data" \
    -v "${redis_config_file}:/usr/local/etc/redis/redis.conf:ro" \
    "${redis_image}" redis-server /usr/local/etc/redis/redis.conf

echo "=== Step 4: Verify ==="
docker exec -e REDISCLI_AUTH="${redis_password}" redis redis-cli ping
unset redis_password
echo "Done. Redis on ${redis_bind_ip}:${redis_port} with persistence enabled and authentication required."

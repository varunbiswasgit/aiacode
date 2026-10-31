#!/usr/bin/env bash
# =============================================================================
# Purpose:      Run Qdrant (vector store) as a Docker container with a
#               persistent volume, an API key, and ports published only on
#               the internal interface.
# Tracker task: P2-07
# Prompts:      qdrant_bind_ip, qdrant_http_port, qdrant_grpc_port,
#               qdrant_volume_name, qdrant_image, qdrant_api_key (secret)
#               qdrant_bind_ip (example — replace with your value): the IP
#               of this node's internal interface, or 127.0.0.1 for a
#               localhost-only deployment. 0.0.0.0 is rejected.
#               qdrant_image (example — replace with your value): a
#               digest-pinned reference, e.g.
#               qdrant/qdrant@sha256:<64 hex digits>. An untagged reference
#               or `latest` is rejected; the operator supplies the digest
#               they verified for the release they intend to run.
# Secrets:      qdrant_api_key is read with `read -s`, written to a
#               mode-600 env file, and passed to the container with
#               --env-file (QDRANT__SERVICE__API_KEY). It is never passed
#               as a script argument, never echoed, and never written to
#               the log file. The verification request sends it via a
#               mode-600 curl config file that is deleted afterwards.
# Prerequisites: Docker installed (02-install-docker.sh). Run ON the DB node.
#               Containers are started by the operator/root, not by the app
#               service account (which has no Docker access).
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-qdrant.log"
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
        echo "A digest-pinned image reference is required, e.g. qdrant/qdrant@sha256:<64 hex digits>. Untagged or 'latest' references are not accepted."
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
        else
            printf '%s' "${first}"; return
        fi
    done
}

echo "=== Step 1: Values ==="
qdrant_bind_ip="$(prompt_bind_ip "Qdrant bind address (internal interface IP, or 127.0.0.1 for localhost-only; never 0.0.0.0)")"
qdrant_http_port="$(prompt "Qdrant HTTP port" "6333")"
qdrant_grpc_port="$(prompt "Qdrant gRPC port" "6334")"
qdrant_volume_name="$(prompt "Docker volume name" "qdrant_storage")"
qdrant_image="$(prompt_image "Qdrant image (example — replace with your value; digest-pinned, e.g. qdrant/qdrant@sha256:<64 hex digits>)")"
qdrant_api_key="$(prompt_secret "Qdrant API key")"

echo "=== Step 2: Store the API key (mode 600) ==="
qdrant_env_file="/etc/airag/qdrant.env"
sudo mkdir -p /etc/airag
sudo chmod 750 /etc/airag
printf 'QDRANT__SERVICE__API_KEY=%s\n' "${qdrant_api_key}" | sudo tee "${qdrant_env_file}" > /dev/null
sudo chmod 600 "${qdrant_env_file}"
echo "API key stored in ${qdrant_env_file} (mode 600)."

echo "=== Step 3: Start Qdrant (idempotent; existing container is recreated) ==="
docker pull "${qdrant_image}"
if docker ps -a --format '{{.Names}}' | grep -qx qdrant; then
    # Recreate, rather than merely restarting, so a pre-existing container
    # without an API key, a pinned digest, or an internal-only bind cannot
    # survive this run. The named volume (data) is preserved.
    docker rm -f qdrant
fi
docker run -d --name qdrant --restart unless-stopped \
    --env-file "${qdrant_env_file}" \
    -p "${qdrant_bind_ip}:${qdrant_http_port}:6333" \
    -p "${qdrant_bind_ip}:${qdrant_grpc_port}:6334" \
    -v "${qdrant_volume_name}:/qdrant/storage" \
    "${qdrant_image}"

echo "=== Step 4: Verify ==="
sleep 3
curl_config="$(mktemp)"
chmod 600 "${curl_config}"
trap 'rm -f "${curl_config}"' EXIT
printf 'header = "Api-Key: %s"\n' "${qdrant_api_key}" > "${curl_config}"
curl -fsS --config "${curl_config}" "http://${qdrant_bind_ip}:${qdrant_http_port}/" | head -c 300; echo
rm -f "${curl_config}"
trap - EXIT
unset qdrant_api_key
docker ps --filter name=qdrant --format '{{.Names}} {{.Status}}'
echo "Done. Qdrant on ${qdrant_bind_ip}:${qdrant_http_port} (HTTP), API key required."

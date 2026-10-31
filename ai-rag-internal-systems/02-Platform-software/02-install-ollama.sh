#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Ollama (local LLM server) on the GPU node.
# Tracker task: P2-09
# Prompts:      ollama_bind_ip, ollama_port, ollama_models_dir,
#               ollama_installer_sha256 (only when Ollama is not already
#               installed)
# Prerequisites: GPU node with working NVIDIA driver (01 script), sudo.
#               Run ON the GPU node. Model download is 06-install-language-model.sh.
# Security:
#   - The installer is NEVER piped to a shell. It is downloaded to a
#     local file, its SHA256 is checked against the expected value the
#     operator supplies from Ollama's published release/download page,
#     and only the local, verified copy is run. A blank or mismatching
#     hash aborts the install; nothing is executed.
#   - Limitation, stated honestly: Ollama publishes no stable checksum
#     file for install.sh itself (its GitHub releases publish a
#     sha256sum.txt for the versioned binary archives, not for the
#     installer script). The expected hash here must therefore be
#     confirmed out of band from the published release/download page
#     before this script runs. A hash fetched from the same host as the
#     script is not independent trust in that host. If no published
#     hash can be confirmed, do not run this installer; obtain a
#     versioned release archive and its published sha256sum.txt instead.
#   - The Ollama API has NO built-in authentication and serves over
#     cleartext HTTP. Binding to an internal interface only is
#     necessary, not sufficient: model traffic must never cross a
#     network where it can be read or where the unauthenticated
#     endpoint can be reached by anyone but the internal services
#     that need it. Never use 0.0.0.0, never expose this port beyond
#     the internal network, and if any client other than the GPU node
#     itself must reach Ollama, put it behind an authenticated TLS
#     reverse proxy (API key and/or SSO) first — binding alone does
#     not authenticate anything.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-ollama.log"
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

prompt_sha256() { # prompt_sha256 <label> — 64 hex digits, never blank
    local value
    while true; do
        read -rp "$1: " value
        if [[ "${value}" =~ ^[0-9a-fA-F]{64}$ ]]; then
            echo "${value}"; return
        fi
        echo "Enter the 64-hex-digit SHA256 published for this installer. A blank or malformed value aborts the install."
    done
}

echo "=== Step 1: Values ==="
ollama_bind_ip="$(prompt_bind_ip "Ollama bind address (internal interface IP, or 127.0.0.1 for localhost-only; never 0.0.0.0)")"
ollama_port="$(prompt "Ollama API port" "11434")"
ollama_models_dir="$(prompt "Model storage directory (example — replace with your value)" "/usr/share/ollama/.ollama/models")"

echo "=== Step 2: Install (idempotent) ==="
if command -v ollama &>/dev/null; then
    echo "Ollama already installed: $(ollama --version)"
else
    ollama_installer_sha256="$(prompt_sha256 "Expected SHA256 of the Ollama installer, from Ollama's published release/download page")"
    installer_dir="$(mktemp -d)"
    trap 'rm -rf "${installer_dir}"' EXIT
    installer_path="${installer_dir}/install.sh"
    # Download only. Nothing is executed from the network.
    curl -fsSL https://ollama.com/install.sh -o "${installer_path}"
    chmod 600 "${installer_path}"
    actual_sha256="$(sha256sum "${installer_path}" | awk '{print $1}')"
    if [ "$(printf '%s' "${actual_sha256}" | tr '[:upper:]' '[:lower:]')" != \
         "$(printf '%s' "${ollama_installer_sha256}" | tr '[:upper:]' '[:lower:]')" ]; then
        echo "ERROR: installer SHA256 mismatch. Expected ${ollama_installer_sha256}, got ${actual_sha256}."
        echo "The installer was NOT run. Check the published value and re-run."
        exit 1
    fi
    echo "Installer SHA256 verified: ${actual_sha256}"
    echo "Downloaded installer: ${installer_path}"
    echo "Inspect it before it runs; it is executed only after your confirmation below."
    confirm=""
    read -rp "Type yes to run the LOCAL, checksum-verified installer: " confirm
    if [ "${confirm}" != "yes" ]; then
        echo "Install aborted by operator. The installer was NOT run."
        exit 1
    fi
    # Run the LOCAL copy only.
    sh "${installer_path}"
fi

echo "=== Step 3: Configure and start ==="
sudo mkdir -p "$(dirname "${ollama_models_dir}")"
sudo systemctl set-environment OLLAMA_HOST="${ollama_bind_ip}:${ollama_port}" 2>/dev/null || true
sudo systemctl enable --now ollama
sleep 2

echo "=== Step 4: Verify ==="
ollama --version
curl -fsS "http://${ollama_bind_ip}:${ollama_port}/api/tags"
echo
echo "Done. Ollama listening on ${ollama_bind_ip}:${ollama_port} (internal interface only)."
echo "Reminder: this endpoint has no authentication — keep it inside the internal network,"
echo "and front it with an authenticated TLS proxy before any non-local client uses it."
echo "Next: 06-install-language-model.sh."

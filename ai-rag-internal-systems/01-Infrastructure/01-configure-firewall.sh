#!/usr/bin/env bash
# =============================================================================
# Purpose:      Configure host firewall (ufw) for one AI-RAG node.
# Tracker task: P1-11
# Prompts:      node_role (app/db/gpu), app_node_ip, db_node_ip, gpu_node_ip,
#               jump_host_ip, siem_ip
# Prerequisites: Ubuntu 24.04, sudo, ufw installed. Run ON each node, choosing
#               that node's role when prompted.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-firewall.log"
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
prompt_optional() { # prompt_optional <label> <default, may be empty>
    local value
    read -rp "$1 [$2]: " value
    echo "${value:-$2}"
}

echo "=== Step 1: Topology ==="
node_role="$(prompt "This node's role (app/db/gpu)" "app")"
app_node_ip="$(prompt "App node IP" "10.0.0.21")"
db_node_ip="$(prompt "DB node IP" "10.0.0.22")"
gpu_node_ip="$(prompt "GPU node IP" "10.0.0.23")"
jump_host_ip="$(prompt "Jump host / admin IP allowed for SSH" "10.0.0.10")"
siem_ip="$(prompt_optional "SIEM IP for log forwarding (blank to skip)" "")"

echo "=== Step 2: Apply rules ==="
sudo apt install -y ufw
sudo ufw --force reset
sudo ufw default deny incoming
sudo ufw default allow outgoing
sudo ufw allow from "${jump_host_ip}" to any port 22 comment 'SSH from jump host'

case "${node_role}" in
    app)
        sudo ufw allow 443/tcp comment 'HTTPS app'
        sudo ufw allow from "${db_node_ip}" to any port 5432 comment 'Postgres (via DB node rules)'
        ;;
    db)
        sudo ufw allow from "${app_node_ip}" to any port 5432 comment 'Postgres from app'
        sudo ufw allow from "${app_node_ip}" to any port 6333 comment 'Qdrant from app'
        sudo ufw allow from "${app_node_ip}" to any port 6379 comment 'Redis from app'
        ;;
    gpu)
        sudo ufw allow from "${app_node_ip}" to any port 11434 comment 'Ollama from app'
        sudo ufw allow from "${app_node_ip}" to any port 8000 comment 'vLLM from app'
        ;;
    *) echo "ERROR: role must be app, db or gpu."; exit 1;;
esac

if [ -n "${siem_ip}" ]; then
    sudo ufw allow out to "${siem_ip}" port 514 comment 'Syslog to SIEM'
fi
sudo ufw --force enable

echo "=== Step 3: Verify ==="
sudo ufw status verbose
echo "Done. Firewall active for role '${node_role}'."

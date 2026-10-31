#!/usr/bin/env bash
# =============================================================================
# Purpose:      Configure host firewall (ufw) for one AI-RAG node.
# Tracker task: P1-11
# Prompts:      node_role (app/db/gpu), app_node_ip, db_node_ip, gpu_node_ip,
#               jump_host_ip, internal_https_ranges, dns_server_ip,
#               ntp_server, proxy_server (optional), proxy_port (only if a
#               proxy is given), apt_mirror (optional), apt_mirror_port
#               (only if a mirror is given), siem_ip (optional), siem_port
#               (only if a SIEM IP is given)
# Prerequisites: Ubuntu 24.04, sudo, ufw installed. Run ON each node, choosing
#               that node's role when prompted.
# Egress model: default deny outgoing. Outbound traffic is allowed ONLY to
#               the explicitly prompted destinations below (proxy, SIEM, DNS,
#               NTP, internal apt mirror) plus, on the app node only, the
#               role-scoped peer service ports the app initiates (DB and GPU
#               nodes). No other outbound destination is allowed.
# =============================================================================
set -euo pipefail

umask 027
LOG_FILE="/var/log/airag/airag-firewall.log"
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
prompt_optional() { # prompt_optional <label> <default, may be empty>
    local value
    read -rp "$1 [$2]: " value
    echo "${value:-$2}"
}

echo "=== Step 1: Topology ==="
node_role="$(prompt "This node's role (app/db/gpu)" "app")"
app_node_ip="$(prompt "App node IP (example — replace with your value)" "10.0.0.21")"
db_node_ip="$(prompt "DB node IP (example — replace with your value)" "10.0.0.22")"
gpu_node_ip="$(prompt "GPU node IP (example — replace with your value)" "10.0.0.23")"
jump_host_ip="$(prompt "Jump host / admin IP allowed for SSH (example — replace with your value)" "10.0.0.10")"
internal_https_ranges="$(prompt "Internal/VPN CIDR range(s) allowed for HTTPS, comma-separated (example — replace with your value)" "10.0.0.0/24")"
dns_server_ip="$(prompt "DNS server IP (example — replace with your value)" "10.0.0.53")"
ntp_server="$(prompt "NTP server IP or hostname (example — replace with your value)" "10.0.0.54")"
proxy_server="$(prompt_optional "Outbound proxy IP or hostname (blank to skip) (example — replace with your value)" "")"
proxy_port=""
if [ -n "${proxy_server}" ]; then
    proxy_port="$(prompt "Outbound proxy port (example — replace with your value)" "3128")"
fi
apt_mirror="$(prompt_optional "Internal apt mirror IP or hostname (blank to skip) (example — replace with your value)" "")"
apt_mirror_port=""
if [ -n "${apt_mirror}" ]; then
    apt_mirror_port="$(prompt "Internal apt mirror port (example — replace with your value)" "443")"
fi
siem_ip="$(prompt_optional "SIEM IP for log forwarding (blank to skip) (example — replace with your value)" "")"
siem_port=""
if [ -n "${siem_ip}" ]; then
    siem_port="$(prompt "SIEM syslog port (example — replace with your value)" "514")"
fi

echo "=== Step 2: Confirm reset and back up current rules ==="
sudo apt install -y ufw
echo "WARNING: The next step will wipe ALL current ufw rules on this node."
echo "Current rules (to be replaced):"
sudo ufw status verbose || true
reset_confirm=""
read -rp "Type YES exactly to confirm wiping the current rules and continue (anything else aborts): " reset_confirm
if [ "${reset_confirm}" != "YES" ]; then
    echo "Aborted by operator: current firewall rules left unchanged."
    exit 1
fi
BACKUP_FILE="/var/log/airag/ufw-rules-backup-$(date +%Y%m%d-%H%M%S).txt"
sudo touch "${BACKUP_FILE}"
sudo chmod 640 "${BACKUP_FILE}"
sudo ufw status verbose | sudo tee "${BACKUP_FILE}" >/dev/null || true
sudo chmod 640 "${BACKUP_FILE}"
echo "Current rules backed up to ${BACKUP_FILE} before reset."

echo "=== Step 3: Apply rules ==="
sudo ufw --force reset
sudo ufw default deny incoming
sudo ufw default deny outgoing
sudo ufw allow from "${jump_host_ip}" to any port 22 comment 'SSH from jump host'

# Egress allowlist (all destinations prompted above; nothing else leaves).
sudo ufw allow out to "${dns_server_ip}" port 53 comment 'DNS outbound'
sudo ufw allow out to "${ntp_server}" port 123 proto udp comment 'NTP outbound'
if [ -n "${proxy_server}" ]; then
    sudo ufw allow out to "${proxy_server}" port "${proxy_port}" proto tcp comment 'Outbound proxy'
fi
if [ -n "${apt_mirror}" ]; then
    sudo ufw allow out to "${apt_mirror}" port "${apt_mirror_port}" proto tcp comment 'Internal apt mirror'
fi
if [ -n "${siem_ip}" ]; then
    sudo ufw allow out to "${siem_ip}" port "${siem_port}" comment 'Syslog to SIEM'
fi

case "${node_role}" in
    app)
        # HTTPS only from the prompted internal/VPN ranges, never any source.
        IFS=',' read -ra https_ranges <<< "${internal_https_ranges}"
        for https_range in "${https_ranges[@]}"; do
            https_range="$(echo "${https_range}" | tr -d '[:space:]')"
            if [ -n "${https_range}" ]; then
                sudo ufw allow from "${https_range}" to any port 443/tcp comment 'HTTPS app (internal/VPN only)'
            fi
        done
        # Outbound peer ports the app initiates. No inbound Postgres rule
        # here: connections originate app -> DB, so only the DB node opens
        # inbound 5432 (see the db case below).
        sudo ufw allow out to "${db_node_ip}" port 5432 proto tcp comment 'Postgres to DB node'
        sudo ufw allow out to "${db_node_ip}" port 6333 proto tcp comment 'Qdrant to DB node'
        sudo ufw allow out to "${db_node_ip}" port 6379 proto tcp comment 'Redis to DB node'
        sudo ufw allow out to "${gpu_node_ip}" port 11434 proto tcp comment 'Ollama to GPU node'
        sudo ufw allow out to "${gpu_node_ip}" port 8000 proto tcp comment 'vLLM to GPU node'
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

sudo ufw --force enable

echo "=== Step 4: Verify ==="
sudo ufw status verbose
echo "Done. Firewall active for role '${node_role}'."

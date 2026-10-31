#!/usr/bin/env bash
# =============================================================================
# Purpose:      Set timezone, NTP sync, and syslog forwarding to the SIEM.
# Tracker task: P1-18
# Prompts:      system_timezone, ntp_server, siem_hostname_or_ip, siem_port
# Prerequisites: Ubuntu 24.04, sudo. Run ON each node.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-time-and-logging.log"
sudo mkdir -p /var/log/airag
sudo touch "${LOG_FILE}" 2>/dev/null || true
sudo chmod 640 "${LOG_FILE}" 2>/dev/null || true
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
prompt_optional() {
    local value
    read -rp "$1 [$2]: " value
    echo "${value:-$2}"
}

echo "=== Step 1: Time and logging values ==="
system_timezone="$(prompt "Timezone (example — replace with your value)" "America/New_York")"
ntp_server="$(prompt "NTP server (example — replace with your value)" "pool.ntp.org")"
siem_hostname_or_ip="$(prompt_optional "SIEM hostname or IP (blank to skip forwarding)" "")"
siem_port="514"
if [ -n "${siem_hostname_or_ip}" ]; then
    siem_port="$(prompt "SIEM syslog port (example — replace with your value)" "514")"
fi

echo "=== Step 2: Timezone and NTP ==="
sudo timedatectl set-timezone "${system_timezone}"
sudo apt install -y systemd-timesyncd
sudo mkdir -p /etc/systemd/timesyncd.conf.d
sudo tee /etc/systemd/timesyncd.conf.d/airag.conf > /dev/null <<EOF
[Time]
NTP=${ntp_server}
EOF
sudo systemctl enable --now systemd-timesyncd
sudo systemctl restart systemd-timesyncd
timedatectl

echo "=== Step 3: Syslog forwarding ==="
if [ -n "${siem_hostname_or_ip}" ]; then
    echo "*.* @${siem_hostname_or_ip}:${siem_port}" | sudo tee /etc/rsyslog.d/50-airag-siem.conf > /dev/null
    sudo systemctl restart rsyslog
    logger "AI-RAG syslog forwarding test from $(hostname)"
    echo "Test log sent to ${siem_hostname_or_ip}:${siem_port} - confirm it arrives in the SIEM."
else
    echo "SIEM forwarding skipped."
fi
echo "Done."

#!/usr/bin/env bash
# =============================================================================
# Purpose:      Forward container and system logs to the SIEM / Graylog.
# Tracker task: P2-13
# Prompts:      siem_hostname_or_ip, siem_syslog_port, docker_log_driver
# Prerequisites: Ubuntu 24.04, sudo, Docker installed. Run ON each node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-log-aggregation.log"
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
siem_hostname_or_ip="$(prompt "SIEM / Graylog hostname or IP" "siem.internal")"
siem_syslog_port="$(prompt "SIEM syslog port" "514")"
docker_log_driver="$(prompt "Docker log driver (syslog/journald/json-file)" "syslog")"

echo "=== Step 2: System logs ==="
echo "*.* @${siem_hostname_or_ip}:${siem_syslog_port}" | sudo tee /etc/rsyslog.d/50-airag-siem.conf > /dev/null
sudo systemctl restart rsyslog

echo "=== Step 3: Docker logs ==="
sudo mkdir -p /etc/docker
sudo tee /etc/docker/daemon.json > /dev/null <<EOF
{
  "log-driver": "${docker_log_driver}",
  "log-opts": {
    "syslog-address": "udp://${siem_hostname_or_ip}:${siem_syslog_port}",
    "tag": "airag-{{.Name}}"
  }
}
EOF
sudo systemctl restart docker

echo "=== Step 4: Verify ==="
logger "AI-RAG log aggregation test from $(hostname)"
docker run --rm hello-world > /dev/null
echo "Test logs sent. Confirm both arrive in the SIEM search before continuing."
echo "Done."

#!/usr/bin/env bash
# =============================================================================
# Purpose:      Set a node's hostname and static IP (netplan) and verify DNS.
# Tracker tasks: P1-05 (name the servers), P1-06 (static IPs and DNS)
# Prompts:      node_hostname, static_ip_address, netmask_prefix, gateway_ip,
#               dns_server_ip, network_interface
# Prerequisites: Ubuntu 24.04, sudo. Run ON the node being configured.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-configure-node-identity.log"
sudo mkdir -p /var/log/airag
sudo touch "${LOG_FILE}" 2>/dev/null || true
sudo chmod 640 "${LOG_FILE}" 2>/dev/null || true
exec > >(sudo tee -a "${LOG_FILE}" >/dev/null) 2>&1 || exec >>"${LOG_FILE}" 2>&1

prompt() { # prompt <label> <default>
    local value
    while true; do
        read -rp "$1 [$2]: " value
        value="${value:-$2}"
        if [ -n "${value}" ]; then echo "${value}"; return; fi
        echo "A value is required. Please try again."
    done
}

echo "=== Step 1: Identity values ==="
node_hostname="$(prompt "Hostname for this node (example — replace with your value)" "PRD-AIRAG-APP01")"
static_ip_address="$(prompt "Static IP address (example — replace with your value)" "10.0.0.21")"
netmask_prefix="$(prompt "Netmask prefix length (example — replace with your value)" "24")"
gateway_ip="$(prompt "Gateway IP (example — replace with your value)" "10.0.0.1")"
dns_server_ip="$(prompt "DNS server IP (example — replace with your value)" "10.0.0.2")"
network_interface="$(prompt "Network interface name (example — replace with your value)" "ens160")"

echo "=== Step 2: Set the hostname ==="
if [ "$(hostname)" = "${node_hostname}" ]; then
    echo "Hostname already ${node_hostname}."
else
    sudo hostnamectl set-hostname "${node_hostname}"
    echo "Hostname set to ${node_hostname}."
fi

echo "=== Step 3: Write the netplan configuration ==="
netplan_file="/etc/netplan/60-airag-static.yaml"
sudo tee "${netplan_file}" > /dev/null <<EOF
network:
  version: 2
  ethernets:
    ${network_interface}:
      dhcp4: false
      addresses: [${static_ip_address}/${netmask_prefix}]
      routes:
        - to: default
          via: ${gateway_ip}
      nameservers:
        addresses: [${dns_server_ip}]
EOF
sudo netplan apply
sleep 2

echo "=== Step 4: Verify ==="
hostname -f
ip -brief address show "${network_interface}"
nslookup "${node_hostname}" || echo "WARNING: DNS lookup failed - create the DNS record and re-run nslookup."
echo "Done. Reboot recommended: sudo reboot"

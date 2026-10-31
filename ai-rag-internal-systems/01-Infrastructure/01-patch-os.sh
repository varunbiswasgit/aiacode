#!/usr/bin/env bash
# =============================================================================
# Purpose:      Patch the OS and enable automatic security updates.
# Tracker task: P1-08
# Prompts:      reboot_after_patch (yes/no)
# Prerequisites: Ubuntu 24.04, sudo, internet or internal apt mirror.
#               Run ON each node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-patch-os.log"
sudo mkdir -p /var/log/airag
exec > >(sudo tee -a "${LOG_FILE}" >/dev/null) 2>&1 || true

prompt() { # prompt <label> <default>
    local value
    read -rp "$1 [$2]: " value
    echo "${value:-$2}"
}

echo "=== Step 1: Patch ==="
sudo apt update
sudo apt upgrade -y
sudo apt autoremove -y

echo "=== Step 2: Enable unattended security updates ==="
sudo apt install -y unattended-upgrades apt-listchanges
echo 'APT::Periodic::Update-Package-Lists "1";' | sudo tee /etc/apt/apt.conf.d/20auto-upgrades > /dev/null
echo 'APT::Periodic::Unattended-Upgrade "1";' | sudo tee -a /etc/apt/apt.conf.d/20auto-upgrades > /dev/null
sudo dpkg-reconfigure -plow unattended-upgrades || true
sudo systemctl enable --now unattended-upgrades 2>/dev/null || true

echo "=== Step 3: Reboot if required ==="
if [ -f /var/run/reboot-required ]; then
    reboot_after_patch="$(prompt "Reboot now (required by updates)? (yes/no)" "yes")"
    if [ "${reboot_after_patch}" = "yes" ]; then
        echo "Rebooting..."
        sudo reboot
    else
        echo "Reboot deferred. Remember to reboot before proceeding."
    fi
else
    echo "No reboot required."
fi
echo "Done. OS patched and auto security updates enabled."

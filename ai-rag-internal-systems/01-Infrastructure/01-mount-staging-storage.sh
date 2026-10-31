#!/usr/bin/env bash
# =============================================================================
# Purpose:      Mount the shared staging storage for document intake.
# Tracker task: P1-14
# Prompts:      nas_hostname_or_ip, nas_export_path, mount_point,
#               mount_owner_account
# Prerequisites: Ubuntu 24.04, sudo, NFS export already shared by the NAS team.
#               Run ON the app node and the DB node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-staging-storage.log"
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

echo "=== Step 1: Share details ==="
nas_hostname_or_ip="$(prompt "NAS hostname or IP" "nas.internal")"
nas_export_path="$(prompt "NAS export path" "/staging")"
mount_point="$(prompt "Local mount point" "/data/staging")"
mount_owner_account="$(prompt "Owner account for the mount" "svc-airag-ingest")"

echo "=== Step 2: Mount ==="
sudo apt install -y nfs-common
sudo mkdir -p "${mount_point}"
fstab_entry="${nas_hostname_or_ip}:${nas_export_path} ${mount_point} nfs defaults,_netdev 0 0"
if grep -qs "${mount_point}" /etc/fstab; then
    echo "fstab entry already present."
else
    echo "${fstab_entry}" | sudo tee -a /etc/fstab > /dev/null
fi
sudo mount "${mount_point}" || sudo mount -a
sudo chown "${mount_owner_account}:${mount_owner_account}" "${mount_point}"

echo "=== Step 3: Verify ==="
df -h "${mount_point}"
sudo -u "${mount_owner_account}" touch "${mount_point}/.write-test" && sudo -u "${mount_owner_account}" rm "${mount_point}/.write-test" \
    && echo "Write test OK for ${mount_owner_account}."
echo "Done. Staging storage mounted at ${mount_point}."

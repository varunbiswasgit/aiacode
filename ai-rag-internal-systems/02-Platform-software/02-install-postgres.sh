#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install PostgreSQL 16 on the data node.
# Tracker task: P2-04
# Prompts:      postgres_version, postgres_data_dir
# Prerequisites: Ubuntu 24.04, sudo, 1 TB data disk available.
#               Run ON the DB node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-postgres.log"
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
postgres_version="$(prompt "PostgreSQL major version" "16")"
postgres_data_dir="$(prompt "Data directory (must exist as a mount for large disks)" "/var/lib/postgresql")"

echo "=== Step 2: Install ==="
if command -v "psql" &>/dev/null && psql --version | grep -q "${postgres_version}"; then
    echo "PostgreSQL ${postgres_version} already installed."
else
    sudo apt update
    sudo apt install -y "postgresql-${postgres_version}"
fi
sudo systemctl enable --now postgresql

echo "=== Step 3: Verify ==="
sudo -u postgres psql -c 'SELECT version();'
echo "Data directory: ${postgres_data_dir} (default cluster)."
echo "Done. Next: 02-enable-pgvector.sh, then 02-create-databases.sql."

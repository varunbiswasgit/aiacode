#!/usr/bin/env bash
# =============================================================================
# Purpose:      Enable the pgvector extension in PostgreSQL.
# Tracker task: P2-05
# Prompts:      postgres_version, target_database
# Prerequisites: PostgreSQL installed (02-install-postgres.sh). Run ON DB node.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-pgvector.log"
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

echo "=== Step 1: Values ==="
postgres_version="$(prompt "PostgreSQL major version" "16")"
target_database="$(prompt "Database to enable pgvector in (created by 02-create-databases.sql; use postgres if not yet created) (example — replace with your value)" "postgres")"

echo "=== Step 2: Install pgvector ==="
sudo apt update
sudo apt install -y "postgresql-${postgres_version}-pgvector"

echo "=== Step 3: Enable and verify ==="
sudo -u postgres psql -d "${target_database}" -c 'CREATE EXTENSION IF NOT EXISTS vector;'
sudo -u postgres psql -d "${target_database}" -c '\dx' | grep -i vector && echo "pgvector enabled in ${target_database}."
echo "Done. Do NOT repeat this extension step manually for the docs or standards databases: 02-create-databases.sql already runs CREATE EXTENSION IF NOT EXISTS vector in both of those. This script's extension step is only needed for a database created outside that script (for example a test database). Note: neither this script nor 02-create-databases.sql enables vector in the audit database — enable it there only if a later phase requires vector in it."

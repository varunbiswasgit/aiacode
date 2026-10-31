#!/usr/bin/env bash
# =============================================================================
# Purpose:      Enable the pgvector extension in PostgreSQL.
# Tracker task: P2-05
# Prompts:      postgres_version, target_database
# Prerequisites: PostgreSQL installed (02-install-postgres.sh). Run ON DB node.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-pgvector.log"
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

postgres_version="$(prompt "PostgreSQL major version" "16")"
target_database="$(prompt "Database to enable pgvector in (created by 02-create-databases.sql; use postgres if not yet created)" "postgres")"

sudo apt update
sudo apt install -y "postgresql-${postgres_version}-pgvector"
sudo -u postgres psql -d "${target_database}" -c 'CREATE EXTENSION IF NOT EXISTS vector;'
sudo -u postgres psql -d "${target_database}" -c '\dx' | grep -i vector && echo "pgvector enabled in ${target_database}."
echo "Done. Repeat the CREATE EXTENSION step for each AI-RAG database after running 02-create-databases.sql."

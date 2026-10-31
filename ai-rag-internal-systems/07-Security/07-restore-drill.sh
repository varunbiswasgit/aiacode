#!/usr/bin/env bash
# =============================================================================
# Purpose:      Verify a backup restore: counts in Postgres and Qdrant on the
#               restored copy must match production.
# Tracker task: P7-07 (backup/restore drill)
# Prompts:      restored_db_host, restored_db_port, restored_db_name,
#               restored_db_user, restored_db_password (secret),
#               restored_qdrant_url, prod_doc_count, prod_collection_count
# Prerequisites: A restored DB01 copy on an isolated network (restore done
#               in the backup console first). Run from the admin workstation.
#               The password is never echoed or saved.
# =============================================================================
set -euo pipefail

prompt() { # prompt <label> <default>
    local value
    while true; do
        read -rp "$1 [$2]: " value
        value="${value:-$2}"
        if [ -n "${value}" ]; then echo "${value}"; return; fi
        echo "A value is required. Please try again."
    done
}

echo "=== Step 1: Restored copy details ==="
restored_db_host="$(prompt "Restored DB host" "PRD-AIRAG-DB01-RESTORED")"
restored_db_port="$(prompt "Restored DB port" "5432")"
restored_db_name="$(prompt "Restored database" "airag_docs")"
restored_db_user="$(prompt "Restored DB user" "airag_app")"
read -rsp "Restored DB password (not echoed): " restored_db_password
echo
restored_qdrant_url="$(prompt "Restored Qdrant URL" "http://PRD-AIRAG-DB01-RESTORED:6333")"
prod_doc_count="$(prompt "Production doc count (from production, for comparison)" "0")"
prod_collection_count="$(prompt "Production Qdrant point count (for comparison)" "0")"

echo "=== Step 2: Compare Postgres ==="
export PGPASSWORD="${restored_db_password}"
restored_count="$(psql -h "${restored_db_host}" -p "${restored_db_port}" -U "${restored_db_user}" \
    -d "${restored_db_name}" -t -A -c "SELECT COUNT(*) FROM clauses;" 2>/dev/null || echo "QUERY-FAILED")"
unset PGPASSWORD restored_db_password
echo "Restored clause count: ${restored_count} (production docs reference: ${prod_doc_count})"

echo "=== Step 3: Compare Qdrant ==="
curl -fsS "${restored_qdrant_url}/collections" | head -c 500; echo
echo "Production Qdrant point count reference: ${prod_collection_count}"

echo "=== Step 4: Verdict ==="
echo "If the restored counts match production, the drill PASSES."
echo "Record the result in the ticket, then delete the restored copy."
echo "Done."

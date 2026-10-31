#!/usr/bin/env bash
# =============================================================================
# Purpose:      Grant exactly INSERT and SELECT on the AI-RAG tables to the
#               application user and to any additional database usernames
#               the operator names. No UPDATE or DELETE is granted on any
#               table, so the append-only rules stay enforced:
#                 standards database: clauses, clause_changes
#                 audit database:     audit_log, feedback, review_decisions
#                 suggestion database: suggestions
#               Re-running this script is safe: GRANT of a privilege that
#               is already held is a no-op, and tables that already have
#               their grants are simply granted again.
# Tracker task: P2-15
# Prompts:      db_host, db_port, db_admin_user, db_admin_password (secret),
#               standards_db, audit_db, suggestions_db, app_user,
#               additional_grantees (optional, space-separated)
# Prerequisites: PostgreSQL 16. The table-creation scripts must already
#               have run (02-create-databases.sql, then
#               04-create-clause-registry.sql and 07-create-audit-log.sql).
#               Run from a host that can reach the DB node. The admin user
#               must be a superuser or the owner of the tables, or the
#               GRANT statements will fail.
# Logic:        For each grantee, for each table in turn: check the table
#               exists; if it does not, print a clear warning and skip
#               that table (other tables are still attempted); otherwise
#               issue GRANT INSERT, SELECT. At the end, print a summary of
#               grants made, tables skipped as missing, and any grants
#               that failed; the script exits non-zero if a grant failed.
# Configuration (defaults are examples — replace with your value):
#               db_host:        localhost (example — replace with your value)
#               db_port:        5432 (example — replace with your value)
#               db_admin_user:  postgres (example — replace with your value)
#               standards_db:   airag_standards (example — replace with your value)
#               audit_db:       airag_audit (example — replace with your value)
#               suggestions_db: same as standards_db (example — replace with your value)
#               app_user:       airag_app (example — replace with your value)
#               additional_grantees: none, blank by default
# Security:     The admin password is read silently with read -s into the
#               PGPASSWORD environment variable only. It is never passed
#               as a command-line argument, never echoed, and never written
#               to the log file or any other file; it is unset on exit
#               (trap) and again at the end of the script. Usernames and
#               table names reach SQL only as psql -v variables, quoted
#               by psql as identifiers (:"name"); no user input is ever
#               interpolated into an SQL string.
#               NOTE on the suggestions table: it is created by a separate
#               suggestion step, not by this script. If it does not exist
#               yet when this script runs, the script warns and skips it —
#               grants must be re-run after the suggestion step creates
#               its table, i.e. run this script again at that point.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-grant-permissions.log"
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

echo "=== Step 1: Connection and databases ==="
db_host="$(prompt "Database host" "localhost")"
db_port="$(prompt "Database port" "5432")"
db_admin_user="$(prompt "Database admin user (superuser or table owner)" "postgres")"
read -rsp "Password for ${db_admin_user} (not echoed): " db_admin_password
echo
export PGPASSWORD="${db_admin_password}"
unset db_admin_password
trap 'unset PGPASSWORD' EXIT

standards_db="$(prompt "Standards database" "airag_standards")"
audit_db="$(prompt "Audit database" "airag_audit")"
suggestions_db="$(prompt "Suggestions database" "${standards_db}")"
app_user="$(prompt "Application database user" "airag_app")"
read -rp "Additional database users to grant to, space-separated (blank for none): " additional_grantees

grantees=("${app_user}")
if [ -n "${additional_grantees}" ]; then
    read -ra extra_grantees <<< "${additional_grantees}"
    grantees+=("${extra_grantees[@]}")
fi

missing_tables=()
failed_grants=()

# grant_table <database> <table> <grantee>
# Per-table attempt: a missing table is a warning, never a fatal error,
# so one not-yet-created table cannot block grants on the other tables.
grant_table() {
    local db_name="$1" table_name="$2" grantee="$3"
    local exists rc
    set +e
    exists="$(psql -h "${db_host}" -p "${db_port}" -U "${db_admin_user}" \
        -d "${db_name}" -t -A -v ON_ERROR_STOP=1 -v tbl="${table_name}" \
        -c "SELECT to_regclass('public.' || :'tbl');" 2>/dev/null)"
    rc=$?
    set -e
    if [ "${rc}" -ne 0 ]; then
        echo "WARNING: could not check table ${table_name} in database ${db_name} (database unreachable or admin login failed); skipping grant to ${grantee}."
        failed_grants+=("${db_name}.${table_name} -> ${grantee} (check failed)")
        return 0
    fi
    if [ -z "${exists}" ]; then
        echo "WARNING: table ${table_name} does not exist yet in database ${db_name}; skipping grant to ${grantee}."
        if [ "${table_name}" = "suggestions" ]; then
            echo "  (The suggestions table is created by a separate suggestion step — re-run this script after that step creates its table.)"
        else
            echo "  (Run that table's creation script first, then re-run this script.)"
        fi
        missing_tables+=("${db_name}.${table_name} -> ${grantee}")
        return 0
    fi
    set +e
    psql -h "${db_host}" -p "${db_port}" -U "${db_admin_user}" \
        -d "${db_name}" -v ON_ERROR_STOP=1 -v tbl="${table_name}" -v grantee="${grantee}" \
        -c 'GRANT INSERT, SELECT ON TABLE public.:"tbl" TO :"grantee";' >/dev/null
    rc=$?
    set -e
    if [ "${rc}" -ne 0 ]; then
        echo "WARNING: GRANT INSERT, SELECT on ${table_name} in database ${db_name} to ${grantee} FAILED."
        failed_grants+=("${db_name}.${table_name} -> ${grantee}")
    else
        echo "Granted INSERT, SELECT on ${table_name} (${db_name}) to ${grantee}."
    fi
}

echo "=== Step 2: Table-level grants (INSERT, SELECT only) ==="
for grantee in "${grantees[@]}"; do
    grant_table "${standards_db}" "clauses" "${grantee}"
    grant_table "${standards_db}" "clause_changes" "${grantee}"
    grant_table "${audit_db}" "audit_log" "${grantee}"
    grant_table "${audit_db}" "feedback" "${grantee}"
    grant_table "${audit_db}" "review_decisions" "${grantee}"
    grant_table "${suggestions_db}" "suggestions" "${grantee}"
done

echo "=== Summary ==="
if [ "${#missing_tables[@]}" -gt 0 ]; then
    echo "Tables skipped because they do not exist yet:"
    for entry in "${missing_tables[@]}"; do
        echo "  ${entry}"
    done
    echo "Re-run this script after the missing tables are created (for suggestions: after the suggestion step creates its table)."
fi
if [ "${#failed_grants[@]}" -gt 0 ]; then
    echo "Grants that FAILED (fix the cause and re-run this script):"
    for entry in "${failed_grants[@]}"; do
        echo "  ${entry}"
    done
    unset PGPASSWORD
    trap - EXIT
    echo "Done with errors."
    exit 1
fi
echo "Done. All existing tables now grant INSERT, SELECT only (no UPDATE, no DELETE) to: ${grantees[*]}."
unset PGPASSWORD
trap - EXIT

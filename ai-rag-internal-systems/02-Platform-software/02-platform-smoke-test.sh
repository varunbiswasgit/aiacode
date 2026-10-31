#!/usr/bin/env bash
# =============================================================================
# Purpose:      Smoke-test the whole platform from the app node: containers,
#               Postgres, Qdrant and Redis reachability.
# Tracker task: P2-14
# Prompts:      db_node_hostname, qdrant_port, redis_port, postgres_port,
#               qdrant_api_key (optional, blank = try without a key)
# Prerequisites: Phases 1-2 complete. Run ON the app node.
# =============================================================================
set -euo pipefail
umask 027

prompt() { # prompt <label> <default>
    local value
    while true; do
        read -rp "$1 [$2]: " value
        value="${value:-$2}"
        if [ -n "${value}" ]; then echo "${value}"; return; fi
        echo "A value is required. Please try again."
    done
}

db_node_hostname="$(prompt "DB node hostname (example — replace with your value)" "PRD-AIRAG-DB01")"
qdrant_port="$(prompt "Qdrant HTTP port (example — replace with your value)" "6333")"
redis_port="$(prompt "Redis port (example — replace with your value)" "6379")"
postgres_port="$(prompt "Postgres port (example — replace with your value)" "5432")"
read -rsp "Qdrant API key (blank for none, not echoed): " qdrant_api_key
echo

failures=0
check() { # check <label> <command...>
    local label="$1"; shift
    if "$@" > /dev/null 2>&1; then echo "OK   - ${label}"; else echo "FAIL - ${label}"; failures=$((failures + 1)); fi
}

echo "=== Platform smoke test ==="
check "Docker daemon" docker info
# Qdrant may require an API key. When a key is given it is sent as an
# Api-Key header; when no key is given and Qdrant answers 401, that
# means authentication is required, not that Qdrant is unreachable.
if [ -n "${qdrant_api_key}" ]; then
    check "Qdrant reachable" curl -fsS -H "Api-Key: ${qdrant_api_key}" "http://${db_node_hostname}:${qdrant_port}/"
else
    qdrant_http_code="$(curl -sS -o /dev/null -w "%{http_code}" "http://${db_node_hostname}:${qdrant_port}/" 2>/dev/null || echo "000")"
    if [ "${qdrant_http_code}" = "401" ]; then
        echo "FAIL - Qdrant requires authentication (HTTP 401): re-run this check with a Qdrant API key; this is not a reachability failure."
        failures=$((failures + 1))
    elif [ "${qdrant_http_code}" -ge 200 ] && [ "${qdrant_http_code}" -lt 300 ]; then
        echo "OK   - Qdrant reachable"
    else
        echo "FAIL - Qdrant reachable"
        failures=$((failures + 1))
    fi
fi
check "Redis reachable" bash -c "exec 3<>/dev/tcp/${db_node_hostname}/${redis_port}"
check "Postgres reachable" bash -c "exec 3<>/dev/tcp/${db_node_hostname}/${postgres_port}"
command -v psql &>/dev/null && pg_isready -h "${db_node_hostname}" -p "${postgres_port}" || echo "NOTE: pg_isready not installed locally; TCP check above covers reachability."
command -v ollama &>/dev/null && ollama list || echo "NOTE: Ollama is on the GPU node, not here - check it there."

if [ "${failures}" -gt 0 ]; then
    echo "RESULT: ${failures} check(s) failed. Fix them before continuing."
    exit 1
fi
echo "RESULT: all platform checks passed."

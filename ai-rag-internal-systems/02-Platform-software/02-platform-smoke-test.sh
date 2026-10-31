#!/usr/bin/env bash
# =============================================================================
# Purpose:      Smoke-test the whole platform from the app node: containers,
#               Postgres, Qdrant and Redis reachability.
# Tracker task: P2-14
# Prompts:      db_node_hostname, qdrant_port, redis_port, postgres_port
# Prerequisites: Phases 1-2 complete. Run ON the app node.
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

db_node_hostname="$(prompt "DB node hostname" "PRD-AIRAG-DB01")"
qdrant_port="$(prompt "Qdrant HTTP port" "6333")"
redis_port="$(prompt "Redis port" "6379")"
postgres_port="$(prompt "Postgres port" "5432")"

failures=0
check() { # check <label> <command...>
    local label="$1"; shift
    if "$@" > /dev/null 2>&1; then echo "OK   - ${label}"; else echo "FAIL - ${label}"; failures=$((failures + 1)); fi
}

echo "=== Platform smoke test ==="
check "Docker daemon" docker info
check "Qdrant reachable" curl -fsS "http://${db_node_hostname}:${qdrant_port}/"
check "Redis reachable" bash -c "exec 3<>/dev/tcp/${db_node_hostname}/${redis_port}"
check "Postgres reachable" bash -c "exec 3<>/dev/tcp/${db_node_hostname}/${postgres_port}"
command -v psql &>/dev/null && pg_isready -h "${db_node_hostname}" -p "${postgres_port}" || echo "NOTE: pg_isready not installed locally; TCP check above covers reachability."
command -v ollama &>/dev/null && ollama list || echo "NOTE: Ollama is on the GPU node, not here - check it there."

if [ "${failures}" -gt 0 ]; then
    echo "RESULT: ${failures} check(s) failed. Fix them before continuing."
    exit 1
fi
echo "RESULT: all platform checks passed."

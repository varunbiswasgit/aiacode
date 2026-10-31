#!/usr/bin/env bash
# =============================================================================
# Purpose:      Prove the app works with general outbound internet blocked
#               (only the standards proxy may be reachable).
# Tracker task: P7-05 (egress test)
# Prompts:      app_base_url, test_question, allowed_proxy_host,
#               blocked_test_host
# Prerequisites: App running, firewall rules from 01-configure-firewall.sh.
#               Run ON the app node. This script only TESTS - it does not
#               change firewall rules.
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

echo "=== Step 1: Test values ==="
app_base_url="$(prompt "App base URL" "https://localhost")"
test_question="$(prompt "Test question" "What does the cleaning validation SOP require?")"
allowed_proxy_host="$(prompt "Allowed standards host (should resolve)" "ecfr.gov")"
blocked_test_host="$(prompt "External host that must be blocked" "example.com")"

echo "=== Step 2: Allowed host check ==="
if getent hosts "${allowed_proxy_host}" > /dev/null; then
    echo "OK - ${allowed_proxy_host} resolves (proxy path intact)."
else
    echo "WARNING - ${allowed_proxy_host} does not resolve; check the proxy rules."
fi

echo "=== Step 3: Blocked host check ==="
if timeout 5 curl -fsS "https://${blocked_test_host}/" > /dev/null 2>&1; then
    echo "WARNING - ${blocked_test_host} reachable! Outbound is NOT blocked as required."
else
    echo "OK - ${blocked_test_host} unreachable (egress blocked as designed)."
fi

echo "=== Step 4: App Q&A check (must succeed fully offline) ==="
for attempt in 1 2 3; do
    if curl -fsS -X POST "${app_base_url}/ask" \
        -H 'Content-Type: application/json' \
        -d "{\"question\": \"${test_question}\"}" | head -c 400; then
        echo; echo "OK - Q&A attempt ${attempt} succeeded."
        break
    else
        echo "Q&A attempt ${attempt} failed."
    fi
    sleep 2
done
echo "Done. Record the result and check egress logs for proxy-only hosts."

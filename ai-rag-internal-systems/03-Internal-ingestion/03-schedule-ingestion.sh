#!/usr/bin/env bash
# =============================================================================
# Purpose:      Create systemd service + timer for scheduled ingestion runs.
# Tracker task: P3-08
# Prompts:      ingest_command, run_user, full_schedule, incremental_schedule,
#               project_dir, log_dir
# Prerequisites: Ubuntu 24.04, sudo, ingest script in place (03-ingest-corpus.py).
#               Run ON the app node.
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

echo "=== Step 1: Schedule values ==="
ingest_command="$(prompt "Ingest command (full path, e.g. venv python + script)" "/opt/airag/venv/bin/python /opt/airag/ingest-corpus.py")"
run_user="$(prompt "User the job runs as" "svc-airag-ingest")"
project_dir="$(prompt "Project directory" "/opt/airag")"
log_dir="$(prompt "Log directory" "/var/log/airag")"
full_schedule="$(prompt "Full re-index schedule (systemd OnCalendar)" "Sun 02:00")"
incremental_schedule="$(prompt "Incremental schedule (systemd OnCalendar)" "hourly")"

echo "=== Step 2: Write units ==="
sudo mkdir -p "${log_dir}"
sudo tee /etc/systemd/system/airag-ingest.service > /dev/null <<EOF
[Unit]
Description=AI-RAG ingestion run
After=network.target

[Service]
Type=oneshot
User=${run_user}
WorkingDirectory=${project_dir}
ExecStart=${ingest_command}
StandardOutput=append:${log_dir}/ingest.log
StandardError=append:${log_dir}/ingest.log
EOF
sudo tee /etc/systemd/system/airag-ingest-full.timer > /dev/null <<EOF
[Unit]
Description=AI-RAG full re-index

[Timer]
OnCalendar=${full_schedule}
Persistent=true

[Install]
WantedBy=timers.target
EOF
sudo tee /etc/systemd/system/airag-ingest-incremental.timer > /dev/null <<EOF
[Unit]
Description=AI-RAG incremental ingest

[Timer]
OnCalendar=${incremental_schedule}
Persistent=true

[Install]
WantedBy=timers.target
EOF
sudo systemctl daemon-reload
sudo systemctl enable --now airag-ingest-full.timer airag-ingest-incremental.timer

echo "=== Step 3: Verify ==="
systemctl list-timers | grep airag
echo "Done. Full: ${full_schedule}; incremental: ${incremental_schedule}."

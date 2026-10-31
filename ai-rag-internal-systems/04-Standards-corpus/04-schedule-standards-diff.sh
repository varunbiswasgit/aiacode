#!/usr/bin/env bash
# =============================================================================
# Purpose:      Schedule the standards-diff job (fetch + detect) via systemd.
# Tracker task: P4-09 (change detector scheduling)
# Prompts:      diff_command, run_user, project_dir, log_dir, extra_write_dirs,
#               diff_schedule
# Prerequisites: Ubuntu 24.04, sudo, phase-04 fetch/detect scripts in place.
#               Run ON the app node.
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

echo "=== Step 1: Schedule values ==="
diff_command="$(prompt "Diff command (fetch + detect, full path) (example — replace with your value)" "/opt/airag/venv/bin/python /opt/airag/detect-standard-changes.py")"
run_user="$(prompt "User the job runs as (example — replace with your value)" "svc-airag-ingest")"
project_dir="$(prompt "Project directory" "/opt/airag")"
log_dir="$(prompt "Log directory" "/var/log/airag")"
extra_write_dirs="$(prompt "Additional writable directories, space-separated (e.g. the staging root fetched standards are written to) (example — replace with your value)" "/data/staging")"
diff_schedule="$(prompt "Diff schedule (systemd OnCalendar)" "daily 02:00")"

echo "=== Step 2: Write units ==="
sudo mkdir -p "${log_dir}"
if [ -f "${log_dir}/standards-diff.log" ]; then sudo chmod 640 "${log_dir}/standards-diff.log"; fi
sudo tee /etc/systemd/system/airag-standards-diff.service > /dev/null <<EOF
[Unit]
Description=AI-RAG standards change detection
After=network.target

[Service]
Type=oneshot
User=${run_user}
WorkingDirectory=${project_dir}
ExecStart=${diff_command}
StandardOutput=append:${log_dir}/standards-diff.log
StandardError=append:${log_dir}/standards-diff.log
UMask=0027
NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=strict
ReadWritePaths=${project_dir} ${log_dir} ${extra_write_dirs}
EOF
sudo tee /etc/systemd/system/airag-standards-diff.timer > /dev/null <<EOF
[Unit]
Description=AI-RAG standards change detection timer

[Timer]
OnCalendar=${diff_schedule}
Persistent=true

[Install]
WantedBy=timers.target
EOF
sudo systemctl daemon-reload
sudo systemctl enable --now airag-standards-diff.timer

echo "=== Step 3: Verify ==="
systemctl list-timers | grep airag
echo "Done. Standards diff scheduled: ${diff_schedule}."

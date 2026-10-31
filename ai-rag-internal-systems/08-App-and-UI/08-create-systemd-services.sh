#!/usr/bin/env bash
# =============================================================================
# Purpose:      Create systemd services so the whole stack starts on boot
#               and restarts on failure.
# Tracker task: P8-09
# Prompts:      app_command, app_user, project_dir, log_dir,
#               enable_model_service (yes/no), model_service_command
# Prerequisites: Ubuntu 24.04, sudo, app installed (08-app-main.py),
#               model server installed (phase 02/06). Run ON the app node
#               (and the GPU node for the model service part).
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

echo "=== Step 1: Service values ==="
app_command="$(prompt "App start command" "/opt/airag/venv/bin/python /opt/airag/08-app-main.py")"
app_user="$(prompt "User the app runs as" "svc-airag-app")"
project_dir="$(prompt "Project directory" "/opt/airag")"
log_dir="$(prompt "Log directory" "/var/log/airag")"
enable_model_service="$(prompt "Also create a vLLM model service on this node? (yes/no)" "no")"
model_service_command=""
if [ "${enable_model_service}" = "yes" ]; then
    model_service_command="$(prompt "Model service command" "/opt/airag/venv/bin/python -m vllm.entrypoints.openai.api_server --model llama3.1:8b --port 8000")"
fi

echo "=== Step 2: App service ==="
sudo mkdir -p "${log_dir}"
sudo tee /etc/systemd/system/airag-app.service > /dev/null <<EOF
[Unit]
Description=AI-RAG web application
After=network.target

[Service]
Type=simple
User=${app_user}
WorkingDirectory=${project_dir}
ExecStart=${app_command}
Restart=always
RestartSec=5
StandardOutput=append:${log_dir}/app.log
StandardError=append:${log_dir}/app.log

[Install]
WantedBy=multi-user.target
EOF

if [ "${enable_model_service}" = "yes" ]; then
    sudo tee /etc/systemd/system/airag-model.service > /dev/null <<EOF
[Unit]
Description=AI-RAG model server (vLLM)
After=network.target

[Service]
Type=simple
User=${app_user}
WorkingDirectory=${project_dir}
ExecStart=${model_service_command}
Restart=always
RestartSec=10
StandardOutput=append:${log_dir}/model.log
StandardError=append:${log_dir}/model.log

[Install]
WantedBy=multi-user.target
EOF
fi

sudo systemctl daemon-reload
sudo systemctl enable --now airag-app.service
if [ "${enable_model_service}" = "yes" ]; then
    sudo systemctl enable --now airag-model.service
fi

echo "=== Step 3: Verify ==="
systemctl status airag-app.service --no-pager | head -5
echo "Reboot test: reboot this node and confirm airag-app (and airag-model) come back up."
echo "Done."

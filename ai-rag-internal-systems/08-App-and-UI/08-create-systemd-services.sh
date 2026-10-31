#!/usr/bin/env bash
# =============================================================================
# Purpose:      Create systemd services so the whole stack starts on boot
#               and restarts on failure.
# Tracker task: P8-09
# Prompts:      app_command, app_user, project_dir, log_dir,
#               enable_model_service (yes/no), model_service_command,
#               env_file, and one prompt per AIRAG_* variable the app
#               reads at startup (see Step 1b). All values are prompted;
#               nothing is hardcoded to a real deployment — defaults
#               shown below are generic examples only, labelled
#               "(example — replace with your value)" where they are
#               placeholders rather than safe generic values.
# Prerequisites: Ubuntu 24.04, sudo, app installed (08-app-main.py),
#               model server installed (phase 02/06). Run ON the app node
#               (and the GPU node for the model service part).
#               The app user (app_user) must already exist.
#
# Unattended startup — EnvironmentFile approach (chosen over inline
# Environment= lines):
#   The app reads every startup value from the environment and otherwise
#   calls input()/getpass(), which under systemd raises EOFError and,
#   with Restart=always, crash-loops. This script therefore prompts the
#   operator once for every AIRAG_* variable the app prompts for, writes
#   them all to a single EnvironmentFile, and points airag-app.service
#   at that file. Every variable is covered, so the app never prompts
#   under systemd. StandardInput=null is also set, so any unexpected
#   prompt fails fast instead of blocking on a terminal that is not there.
#   Environment= lines were NOT chosen: unit files in
#   /etc/systemd/system are world-readable and their Environment values
#   are visible via systemctl show, so a secret (AIRAG_DB_PASSWORD)
#   must never live there. The EnvironmentFile below is created
#   mode 600 and owned by the app user, and the password is read with
#   read -s, never printed, never passed in argv, and never written to
#   any log. Do not cat the env file or add its contents to a unit.
#
# Hardening (applied to EVERY generated unit):
#   NoNewPrivileges=true, PrivateTmp=true, ProtectSystem=strict,
#   UMask=0027, plus ReadWritePaths for the log directory — the only
#   path these services legitimately write to in this script's scope
#   (logs are written by systemd to the pre-created 0640 log files;
#   the app itself writes no files, and the model server's scratch
#   space is its private /tmp). If a model server command needs an
#   extra writable directory (e.g. a model cache), add that directory
#   to the model unit's ReadWritePaths by hand — do not weaken
#   ProtectSystem to work around it.
#
# Log permissions: this script runs with umask 027, creates the log
#   directory mode 0750 owned by the app user, and pre-creates each
#   log file mode 0640 owned by the app user, so logs are never
#   world-readable.
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

prompt_required() { # prompt_required <label> — no default; an example
                    # placeholder, if shown in the label, must be replaced
    local value
    while true; do
        read -rp "$1: " value
        if [ -n "${value}" ]; then echo "${value}"; return; fi
        echo "A value is required. Please try again."
    done
}

prompt_secret() { # prompt_secret <label> — input is not echoed and the
                  # value is only ever captured into a variable, never
                  # printed, logged, or passed as an argument
    local value
    while true; do
        read -rsp "$1: " value
        echo
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

echo "=== Step 1b: App environment (AIRAG_*) ==="
echo "These values are written to the EnvironmentFile in Step 2 so the"
echo "app starts unattended under systemd without prompting. Answer every"
echo "prompt; placeholders marked as examples must be replaced with your"
echo "real values. The database password is not echoed."
env_file="$(prompt "Environment file path" "/etc/airag/airag-app.env")"
airag_host="$(prompt "Bind host" "0.0.0.0")"
airag_port="$(prompt "Bind port" "8000")"
airag_qdrant_url="$(prompt_required "Qdrant base URL (example — replace with your value: http://<qdrant-host>:6333)")"
airag_collection="$(prompt "Qdrant collection" "airag_docs")"
airag_model_backend="$(prompt "Model backend (ollama/vllm)" "ollama")"
airag_model_api_url="$(prompt_required "Model API URL (example — replace with your value: http://<model-host>:11434)")"
airag_model_name="$(prompt "Model name (example — replace with your value)" "llama3.1:8b")"
airag_db_host="$(prompt_required "Audit database host (example — replace with your value: <db-host>)")"
airag_db_port="$(prompt "Audit database port" "5432")"
airag_db_name="$(prompt "Audit database name" "airag_audit")"
airag_db_user="$(prompt "Audit database user" "airag_app")"
airag_db_password="$(prompt_secret "Audit database password (not echoed)")"
airag_prompts_dir="$(prompt "Prompts directory" "${project_dir}/prompts")"
airag_access_map="$(prompt "Access map JSON path" "${project_dir}/access-map.json")"
airag_app_base_url="$(prompt "App public base URL for OIDC redirect (example — replace with your value)" "http://localhost:8000")"
airag_embedding_model="$(prompt "Embedding model (example — replace with your value)" "BAAI/bge-large-en-v1.5")"
airag_reranker_model="$(prompt "Reranker model (example — replace with your value)" "BAAI/bge-reranker-large")"
airag_final_top_k="$(prompt "Final top-k to keep" "8")"
airag_retrieve_top_k="$(prompt "Candidates per path" "50")"
airag_rrf_k="$(prompt "RRF constant k" "60")"
airag_min_rerank_score="$(prompt "Minimum rerank score to answer" "0.0")"
airag_temperature="$(prompt "Temperature (0.1-0.3 for factual answers)" "0.2")"
airag_audit_db_name="$(prompt "Audit database (example — replace with your value)" "airag_audit")"
airag_standards_db_name="$(prompt "Standards database (example — replace with your value)" "airag_standards")"
airag_docs_db_name="$(prompt "Documents database for keyword search (example — replace with your value)" "airag_docs")"
airag_user_groups_map="$(prompt "Server-side user-groups map JSON path (example — replace with your value)" "${project_dir}/user-groups.json")"
airag_sso_config="$(prompt "SSO config path written by 07-configure-sso.sh (example — replace with your value)" "${project_dir}/sso-config.yaml")"
airag_pipeline_dir="$(prompt "Phase-05 pipeline scripts directory (example — replace with your value)" "${project_dir}/05-RAG-core")"
airag_session_ttl="$(prompt "Session lifetime (seconds)" "3600")"
airag_session_cookie_secure="$(prompt "Mark session cookie Secure (yes/no; yes behind TLS)" "no")"
airag_ldap_user_dn_template="$(prompt "LDAP user DN template with {username} (example — replace with your value; only used when SSO type is LDAP)" "uid={username},ou=users,dc=example,dc=com")"

echo "=== Step 2: Environment file and log files ==="
# Create the env file mode 600 owned by the app user BEFORE writing any
# content into it, so the secret is never briefly world-readable.
# The secret travels only via stdin to tee (never argv, never echoed).
if ! sudo install -d -m 0750 -o root -g "${app_user}" "$(dirname "${env_file}")" 2>/dev/null; then
    # Fallback when no group matches the app user: the app user still
    # owns the directory, so it can traverse to its 0600 env file.
    sudo mkdir -p "$(dirname "${env_file}")"
    sudo chown "${app_user}" "$(dirname "${env_file}")"
    sudo chmod 750 "$(dirname "${env_file}")"
fi
sudo install -m 0600 -o "${app_user}" /dev/null "${env_file}"
sudo tee "${env_file}" > /dev/null <<EOF
AIRAG_HOST=${airag_host}
AIRAG_PORT=${airag_port}
AIRAG_QDRANT_URL=${airag_qdrant_url}
AIRAG_COLLECTION=${airag_collection}
AIRAG_MODEL_BACKEND=${airag_model_backend}
AIRAG_MODEL_API_URL=${airag_model_api_url}
AIRAG_MODEL_NAME=${airag_model_name}
AIRAG_DB_HOST=${airag_db_host}
AIRAG_DB_PORT=${airag_db_port}
AIRAG_DB_NAME=${airag_db_name}
AIRAG_DB_USER=${airag_db_user}
AIRAG_DB_PASSWORD=${airag_db_password}
AIRAG_PROMPTS_DIR=${airag_prompts_dir}
AIRAG_ACCESS_MAP=${airag_access_map}
AIRAG_APP_BASE_URL=${airag_app_base_url}
AIRAG_EMBEDDING_MODEL=${airag_embedding_model}
AIRAG_RERANKER_MODEL=${airag_reranker_model}
AIRAG_FINAL_TOP_K=${airag_final_top_k}
AIRAG_RETRIEVE_TOP_K=${airag_retrieve_top_k}
AIRAG_RRF_K=${airag_rrf_k}
AIRAG_MIN_RERANK_SCORE=${airag_min_rerank_score}
AIRAG_TEMPERATURE=${airag_temperature}
AIRAG_AUDIT_DB_NAME=${airag_audit_db_name}
AIRAG_STANDARDS_DB_NAME=${airag_standards_db_name}
AIRAG_DOCS_DB_NAME=${airag_docs_db_name}
AIRAG_USER_GROUPS_MAP=${airag_user_groups_map}
AIRAG_SSO_CONFIG=${airag_sso_config}
AIRAG_PIPELINE_DIR=${airag_pipeline_dir}
AIRAG_SESSION_TTL=${airag_session_ttl}
AIRAG_SESSION_COOKIE_SECURE=${airag_session_cookie_secure}
AIRAG_LDAP_USER_DN_TEMPLATE=${airag_ldap_user_dn_template}
EOF
sudo chmod 600 "${env_file}"
sudo chown "${app_user}" "${env_file}"
# The password variable is no longer needed; drop it from this shell.
airag_db_password=""
# Log directory and log files: 0750 / 0640, owned by the app user.
sudo install -d -m 0750 -o "${app_user}" "${log_dir}"
sudo install -m 0640 -o "${app_user}" /dev/null "${log_dir}/app.log"
if [ "${enable_model_service}" = "yes" ]; then
    sudo install -m 0640 -o "${app_user}" /dev/null "${log_dir}/model.log"
fi

echo "=== Step 3: App service ==="
sudo tee /etc/systemd/system/airag-app.service > /dev/null <<EOF
[Unit]
Description=AI-RAG web application
After=network.target

[Service]
Type=simple
User=${app_user}
WorkingDirectory=${project_dir}
EnvironmentFile=${env_file}
ExecStart=${app_command}
Restart=always
RestartSec=5
StandardInput=null
StandardOutput=append:${log_dir}/app.log
StandardError=append:${log_dir}/app.log
NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=strict
ReadWritePaths=${log_dir}
UMask=0027

[Install]
WantedBy=multi-user.target
EOF

if [ "${enable_model_service}" = "yes" ]; then
    echo "=== Step 4: Model service ==="
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
StandardInput=null
StandardOutput=append:${log_dir}/model.log
StandardError=append:${log_dir}/model.log
NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=strict
ReadWritePaths=${log_dir}
UMask=0027

[Install]
WantedBy=multi-user.target
EOF
fi

sudo systemctl daemon-reload
sudo systemctl enable --now airag-app.service
if [ "${enable_model_service}" = "yes" ]; then
    sudo systemctl enable --now airag-model.service
fi

echo "=== Step 5: Verify ==="
systemctl status airag-app.service --no-pager | head -5
echo "Reboot test: reboot this node and confirm airag-app (and airag-model) come back up."
echo "Done."

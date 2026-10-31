#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install Tesseract OCR and ocrmypdf for scanned PDFs.
# Tracker task: P3-07
# Prompts:      venv_path, ocr_languages
# Prerequisites: Ubuntu 24.04, sudo, project venv (02-install-python.sh).
#               Run ON the app node.
# =============================================================================
set -euo pipefail
umask 027

LOG_FILE="/var/log/airag-ocr.log"
sudo mkdir -p /var/log/airag
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

venv_path="$(prompt "Virtual environment path (example — replace with your value)" "/opt/airag/venv")"
ocr_languages="$(prompt "OCR languages (tesseract codes, comma-separated) (example — replace with your value)" "eng")"

sudo apt update
sudo apt install -y tesseract-ocr
# Install language packs for each requested language (eng is in the base package).
for language in ${ocr_languages//,/ }; do
    if [ "${language}" != "eng" ]; then
        sudo apt install -y "tesseract-ocr-${language}" || echo "WARNING: pack for ${language} not found, continuing."
    fi
done
"${venv_path}/bin/pip" install ocrmypdf

echo "=== Verify ==="
tesseract --version | head -1
"${venv_path}/bin/ocrmypdf" --version
echo "Done. OCR a scanned PDF with: ocrmypdf --jobs 4 input.pdf output.pdf"

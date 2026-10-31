#!/usr/bin/env bash
# =============================================================================
# Purpose:      Install the NVIDIA driver and CUDA toolkit, verify the GPU.
# Tracker task: P1-13
# Prompts:      install_cuda_toolkit (yes/no), reboot_after_driver (yes/no)
# Prerequisites: Ubuntu 24.04, GPU attached via passthrough (P1-12), sudo.
#               Run ON the GPU node only.
# =============================================================================
set -euo pipefail

LOG_FILE="/var/log/airag-nvidia-driver.log"
sudo mkdir -p /var/log/airag
exec > >(sudo tee -a "${LOG_FILE}" >/dev/null) 2>&1 || true

prompt() { # prompt <label> <default>
    local value
    read -rp "$1 [$2]: " value
    echo "${value:-$2}"
}

echo "=== Step 1: Detect the GPU ==="
lspci | grep -i nvidia || { echo "ERROR: no NVIDIA card visible. Attach it via passthrough first (P1-12)."; exit 1; }

echo "=== Step 2: Install the driver ==="
if command -v nvidia-smi &>/dev/null && nvidia-smi -L &>/dev/null; then
    echo "NVIDIA driver already working:"
    nvidia-smi -L
else
    sudo apt update
    sudo ubuntu-drivers autoinstall
    echo "Driver installed."
fi

reboot_after_driver="$(prompt "Reboot now to load the driver? (yes/no)" "yes")"
if [ "${reboot_after_driver}" = "yes" ]; then
    echo "Rebooting - re-run this script afterwards to continue with CUDA."
    sudo reboot
fi

echo "=== Step 3: Install the CUDA toolkit (optional) ==="
install_cuda_toolkit="$(prompt "Install CUDA toolkit? (yes/no)" "yes")"
if [ "${install_cuda_toolkit}" = "yes" ]; then
    if command -v nvcc &>/dev/null; then
        echo "CUDA toolkit already installed."
    else
        sudo apt install -y nvidia-cuda-toolkit
    fi
    nvcc --version
fi

echo "=== Step 4: Verify ==="
nvidia-smi
echo "Done. GPU ready for Ollama / vLLM."

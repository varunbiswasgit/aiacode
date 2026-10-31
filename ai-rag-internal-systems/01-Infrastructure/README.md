# 01 — Infrastructure

## Purpose

Prepare the three AI-RAG nodes (GPU/inference, app/API, data/DB) on
Ubuntu 24.04: identity and networking, patching, Active Directory join,
service accounts, host firewall, GPU driver, shared storage, monitoring
and time/log forwarding. Console-only steps (slicing VMs in the
hypervisor, attaching the GPU, backup jobs, snapshots) are deliberately
not scripted and stay manual in the tracker.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 01-configure-node-identity.sh | P1-05, P1-06 | Sets the hostname and writes a static-IP netplan config, then verifies DNS. |
| 01-patch-os.sh | P1-08 | Apt upgrade, autoremove, enables unattended security updates, optional reboot. |
| 01-join-ad-domain.sh | P1-09 | Installs realmd/sssd, joins the AD domain, verifies a domain login lookup. |
| 01-create-service-accounts.sh | P1-10 | Creates the app/ingest/db service accounts (locked passwords, SSH-key login) and optionally adds the app account to the docker group. |
| 01-configure-firewall.sh | P1-11 | Resets ufw and applies role-based rules (app, db, gpu) plus jump-host SSH and optional SIEM egress. |
| 01-install-nvidia-driver.sh | P1-13 | Detects the passed-through GPU, installs the driver via ubuntu-drivers, optional CUDA toolkit, verifies with nvidia-smi. |
| 01-mount-staging-storage.sh | P1-14 | Mounts the NFS staging share, persists it in fstab, verifies write access for the ingest account. |
| 01-install-monitoring-agent.sh | P1-17 | Downloads and installs Prometheus node_exporter as a systemd service. |
| 01-configure-time-and-logging.sh | P1-18 | Sets timezone and NTP, optionally forwards syslog to the SIEM. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 01-configure-node-identity.sh | node_hostname, static_ip_address, netmask_prefix, gateway_ip, dns_server_ip, network_interface |
| 01-patch-os.sh | reboot_after_patch (yes/no) |
| 01-join-ad-domain.sh | ad_domain_name, ad_admin_username, ad_admin_password (secret), test_domain_username |
| 01-create-service-accounts.sh | app_service_account, ingest_service_account, db_service_account, add_to_docker_group (yes/no) |
| 01-configure-firewall.sh | node_role (app/db/gpu), app_node_ip, db_node_ip, gpu_node_ip, jump_host_ip, siem_ip (optional) |
| 01-install-nvidia-driver.sh | install_cuda_toolkit (yes/no), reboot_after_driver (yes/no) |
| 01-mount-staging-storage.sh | nas_hostname_or_ip, nas_export_path, mount_point, mount_owner_account |
| 01-install-monitoring-agent.sh | node_exporter_version, install_dir, metrics_port |
| 01-configure-time-and-logging.sh | system_timezone, ntp_server, siem_hostname_or_ip (optional), siem_port |

Secrets (the AD admin password) are read with `read -s`, never echoed,
and cleared from memory after use.

## Logic flow

Run in tracker order on each node: identity → patch → AD join → service
accounts → firewall → (GPU node only) NVIDIA driver → staging mount →
monitoring → time/logging. Every script is idempotent where possible
(already-set values are detected and skipped) and logs to
`/var/log/airag/`.

## Configuration

- Default hostnames: PRD-AIRAG-GPU01, PRD-AIRAG-APP01, PRD-AIRAG-DB01
  (all prompted, defaults shown only as examples).
- Default node_exporter version: 1.8.2, port 9100.
- Default timezone: America/New_York.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

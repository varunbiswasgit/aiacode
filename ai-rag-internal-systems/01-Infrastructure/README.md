AI-RAG Internal Systems is an on-premises retrieval-augmented generation platform for internal content discovery, question answering over internal documents and regulatory standards (FDA eCFR Title 21, EU GMP, ICH), and document-update recommendations when a standard changes. It is generic: every deployment value is supplied as user input at runtime, and no company-specific value is hardcoded. Every recommendation is a suggestion for human QA and document-owner approval; the system never auto-edits a controlled document.

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
| 01-create-service-accounts.sh | P1-10 | Creates the app/ingest/db service accounts (locked passwords, SSH-key login). The docker-group prompt is removed: no account is given Docker access, and existing docker-group membership for the app account is actively removed. |
| 01-configure-firewall.sh | P1-11 | Resets ufw and applies role-based rules (app, db, gpu) plus jump-host SSH. Outbound traffic defaults to DENY; only the prompted allowlist is allowed (proxy, SIEM, DNS, NTP, internal apt mirror) plus, on the app node only, the outbound peer ports the app initiates. Reset requires typing YES and first writes a timestamped backup of the current rules under /var/log/airag/. The app role has no inbound 5432 rule (connections originate app to DB, so only the DB node opens inbound 5432), and 443 is restricted to the prompted internal/VPN ranges. Log: /var/log/airag/airag-firewall.log (mode 640). |
| 01-install-nvidia-driver.sh | P1-13 | Detects the passed-through GPU, installs the driver via ubuntu-drivers, optional CUDA toolkit, verifies with nvidia-smi. |
| 01-mount-staging-storage.sh | P1-14 | Mounts the NFS staging share, persists it in fstab, verifies write access for the ingest account. |
| 01-install-monitoring-agent.sh | P1-17 | Downloads and installs Prometheus node_exporter as a systemd service. Metrics are bound to 127.0.0.1 only, so remote scraping fails by design; scrape from a collector on the same node (or via an SSH tunnel / local agent). |
| 01-configure-time-and-logging.sh | P1-18 | Sets timezone and NTP, optionally forwards syslog to the SIEM. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 01-configure-node-identity.sh | node_hostname, static_ip_address, netmask_prefix, gateway_ip, dns_server_ip, network_interface |
| 01-patch-os.sh | reboot_after_patch (yes/no) |
| 01-join-ad-domain.sh | ad_domain_name, ad_admin_username, ad_admin_password (secret), test_domain_username |
| 01-create-service-accounts.sh | app_service_account, ingest_service_account, db_service_account |
| 01-configure-firewall.sh | node_role (app/db/gpu), app_node_ip, db_node_ip, gpu_node_ip, jump_host_ip, internal_https_ranges, dns_server_ip, ntp_server, proxy_server (optional) and proxy_port (only if a proxy is given), apt_mirror (optional) and apt_mirror_port (only if a mirror is given), siem_ip (optional) and siem_port (only if a SIEM IP is given), plus a YES confirmation before the reset |
| 01-install-nvidia-driver.sh | install_cuda_toolkit (yes/no), reboot_after_driver (yes/no) |
| 01-mount-staging-storage.sh | nas_hostname_or_ip, nas_export_path, mount_point, mount_owner_account |
| 01-install-monitoring-agent.sh | node_exporter_version, install_dir, metrics_port, expected_sha256 (optional: blank means verify against the release's published sha256sums.txt instead) |
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

- Prompt labels carry "(example — replace with your value)"; scripts run
  with umask 027 and log files are mode 640.
- Default hostnames: PRD-AIRAG-GPU01, PRD-AIRAG-APP01, PRD-AIRAG-DB01
  (all prompted, defaults shown only as examples).
- Default node_exporter version: 1.8.2, port 9100. The version is
  validated against ^[0-9]+\.[0-9]+\.[0-9]+$ before it is used in the
  download URL; anything else is refused. An expected_sha256 value, when
  given, must be 64 hex characters and is checked directly against the
  downloaded archive; when blank, the release's published
  sha256sums.txt is used instead.
- Default timezone: America/New_York.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |
| v2 | Fix pass: firewall default-deny egress with prompted allowlist, YES-confirmed reset with timestamped backup under /var/log/airag/, app-role inbound 5432 removed, 443 restricted to prompted internal/VPN ranges, log moved to /var/log/airag/airag-firewall.log (640); monitoring optional expected_sha256, version validated, metrics bound to 127.0.0.1 only; service accounts: docker-group prompt removed and existing membership actively removed; other scripts: labels, umask 027, logs 640. |

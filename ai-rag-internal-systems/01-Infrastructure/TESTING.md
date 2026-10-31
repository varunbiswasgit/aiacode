# Testing — 01 Infrastructure

Manual test cases. Run each case, then tick the Pass column. Use a
disposable test VM for every case that changes system state.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Test VM with DHCP address | Run 01-configure-node-identity.sh with a test hostname, IP, gateway and DNS | hostname -f shows the new name, ip shows the static address, nslookup resolves the name | ☐ |
| Test VM unpatched | Run 01-patch-os.sh, answer no to reboot | apt upgrade completes, unattended-upgrades enabled, no pending package errors | ☐ |
| Test VM not domain-joined, test AD available | Run 01-join-ad-domain.sh with test domain and admin | realm list shows the domain, id for the test user resolves | ☐ |
| Test VM, no service accounts | Run 01-create-service-accounts.sh with test account names (no docker-group prompt is offered) | id shows all three accounts, password login locked, no account is in the docker group; if the app account was already in the docker group, the script removes it and the final check passes only when it is not a member | ☐ |
| Test VM, ufw inactive, existing ufw rules present | Run 01-configure-firewall.sh with role app and test IPs, internal/VPN ranges, DNS, NTP and (optionally) proxy/apt-mirror/SIEM values; type YES at the reset confirmation | A timestamped backup of the previous rules appears under /var/log/airag/ (ufw-rules-backup-YYYYMMDD-HHMMSS.txt); ufw status verbose shows default deny incoming and default deny outgoing, SSH only from the jump host, 443 only from the prompted internal/VPN ranges, no inbound 5432 on the app node, outbound peer rules to the DB node (5432, 6333, 6379) and GPU node (11434, 8000), plus the prompted DNS/NTP/proxy/apt-mirror/SIEM egress rules; log is /var/log/airag/airag-firewall.log (mode 640). Answering anything other than YES aborts with the old rules unchanged | ☐ |
| Test VM with GPU attached via passthrough | Run 01-install-nvidia-driver.sh, answer yes to CUDA | lspci lists the NVIDIA card, nvidia-smi prints the GPU, nvcc --version works | ☐ |
| Test VM, NAS export shared | Run 01-mount-staging-storage.sh with test NAS path and mount point | df -h shows the mount, write test as the ingest account succeeds, fstab entry persists after reboot | ☐ |
| Test VM, no monitoring | Run 01-install-monitoring-agent.sh with test version and port, optionally with an expected_sha256 | node_exporter service active, curl to localhost on the metrics port returns metrics lines; the service listens on 127.0.0.1 only, so a remote scrape of the metrics port fails by design. A non-X.Y.Z version is refused before download; a supplied checksum mismatch aborts the install | ☐ |
| Test VM, default timezone | Run 01-configure-time-and-logging.sh with test timezone, NTP and SIEM | timedatectl shows the timezone and sync, test syslog message arrives at the SIEM | ☐ |

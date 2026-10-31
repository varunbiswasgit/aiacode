# Testing — 01 Infrastructure

Manual test cases. Run each case, then tick the Pass column. Use a
disposable test VM for every case that changes system state.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Test VM with DHCP address | Run 01-configure-node-identity.sh with a test hostname, IP, gateway and DNS | hostname -f shows the new name, ip shows the static address, nslookup resolves the name | ☐ |
| Test VM unpatched | Run 01-patch-os.sh, answer no to reboot | apt upgrade completes, unattended-upgrades enabled, no pending package errors | ☐ |
| Test VM not domain-joined, test AD available | Run 01-join-ad-domain.sh with test domain and admin | realm list shows the domain, id for the test user resolves | ☐ |
| Test VM, no service accounts | Run 01-create-service-accounts.sh with test account names | id shows all three accounts, password login locked, docker group membership applied when requested | ☐ |
| Test VM, ufw inactive | Run 01-configure-firewall.sh with role app and test IPs | ufw status verbose shows SSH only from the jump host plus 443, app rules present | ☐ |
| Test VM with GPU attached via passthrough | Run 01-install-nvidia-driver.sh, answer yes to CUDA | lspci lists the NVIDIA card, nvidia-smi prints the GPU, nvcc --version works | ☐ |
| Test VM, NAS export shared | Run 01-mount-staging-storage.sh with test NAS path and mount point | df -h shows the mount, write test as the ingest account succeeds, fstab entry persists after reboot | ☐ |
| Test VM, no monitoring | Run 01-install-monitoring-agent.sh with test version and port | node_exporter service active, curl to the metrics port returns metrics lines | ☐ |
| Test VM, default timezone | Run 01-configure-time-and-logging.sh with test timezone, NTP and SIEM | timedatectl shows the timezone and sync, test syslog message arrives at the SIEM | ☐ |

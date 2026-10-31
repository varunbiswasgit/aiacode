#!/usr/bin/env bash
# =============================================================================
# Purpose:      Create the pilot-users AD group and add the pilot members.
# Tracker task: P9-04 (select pilot users)
# Prompts:      ad_domain_name, pilot_group_name, group_member_usernames,
#               ad_admin_username
# Prerequisites: Domain-joined admin workstation, sudo, samba-tool available
#               for the scripted path. If samba-tool is absent, the script
#               prints the equivalent Active Directory Users and Computers
#               steps instead of failing.
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

echo "=== Step 1: Group details ==="
ad_domain_name="$(prompt "Active Directory domain (example — replace with your value)" "corp.example.com")"
pilot_group_name="$(prompt "Pilot group name (example — replace with your value)" "AIRAG-Pilot")"
group_member_usernames="$(prompt "Pilot member usernames, comma-separated (example — replace with your value)" "qa.user,ops.user")"
ad_admin_username="$(prompt "Domain admin username (example — replace with your value)" "administrator")"

echo "=== Step 2: Create group and add members ==="
if command -v samba-tool &>/dev/null; then
    if samba-tool group list 2>/dev/null | grep -qx "${pilot_group_name}"; then
        echo "Group ${pilot_group_name} already exists."
    else
        samba-tool group add "${pilot_group_name}" -U "${ad_admin_username}@${ad_domain_name}"
        echo "Group ${pilot_group_name} created."
    fi
    for member in ${group_member_usernames//,/ }; do
        samba-tool group addmembers "${pilot_group_name}" "${member}" \
            -U "${ad_admin_username}@${ad_domain_name}" \
            && echo "Added ${member}." \
            || echo "WARNING: could not add ${member} (check the username)."
    done
else
    echo "samba-tool not found. Do this manually in Active Directory Users and Computers:"
    echo "  1. Create group ${pilot_group_name} in domain ${ad_domain_name}."
    echo "  2. Add members: ${group_member_usernames}."
    read -rp "Press Enter once the group and members exist... "
fi

echo "=== Step 3: Verify ==="
getent group "${pilot_group_name}" || echo "NOTE: group not visible via getent on this host - verify in AD directly."
echo "Done. Pilot group: ${pilot_group_name}."

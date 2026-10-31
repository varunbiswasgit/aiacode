#!/usr/bin/env python3
"""Test read access to one internal document source and fetch a sample.

Tracker tasks: P3-02 (read-only access check), P3-03 (source connectors)
Prompts: source_type, site/share URL, username, access token (secret),
         sample_output_path
Prerequisites: network route to the source; a read-only service account.
  - SharePoint: a Microsoft Graph access token for the service account.
  - Confluence: a Confluence API token (basic auth user + token).
  - SMB share: smbclient installed; the script shells out to it.
Only the Python standard library is used. No secret is printed or saved.
"""
import getpass
import json
import subprocess
import sys
import urllib.request


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_secret(label):
    while True:
        value = getpass.getpass(f"{label}: ").strip()
        if value:
            return value
        print("A value is required.")


def main():
    print("=== Step 1: Source details ===")
    source_type = prompt("Source type (sharepoint/smb/confluence)", "sharepoint").lower()
    source_url = prompt("Source URL or share path (e.g. site URL, //server/share, Confluence base URL)")
    username = prompt("Service account username", "svc-airag-ingest")
    access_token = prompt_secret("Access token / password (not echoed, never saved)")
    sample_output_path = prompt("Where to save the sample listing", "/tmp/airag-source-sample.json")

    print(f"=== Step 2: Testing read access ({source_type}) ===")
    sample = {}
    if source_type == "sharepoint":
        request = urllib.request.Request(
            f"{source_url.rstrip('/')}/_api/web/lists",
            headers={"Authorization": f"Bearer {access_token}",
                     "Accept": "application/json"})
        with urllib.request.urlopen(request, timeout=30) as response:
            sample = json.loads(response.read().decode("utf-8"))
        print(f"SharePoint reachable; lists returned: {len(sample.get('value', []))}")
    elif source_type == "confluence":
        request = urllib.request.Request(
            f"{source_url.rstrip('/')}/rest/api/content/search?cql=type=page&limit=5",
            headers={"Authorization": f"Bearer {access_token}",
                     "Accept": "application/json"})
        with urllib.request.urlopen(request, timeout=30) as response:
            sample = json.loads(response.read().decode("utf-8"))
        print(f"Confluence reachable; pages returned: {len(sample.get('results', []))}")
    elif source_type == "smb":
        result = subprocess.run(
            ["smbclient", source_url, "-U", f"{username}%{access_token}", "-c", "ls"],
            capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            print("SMB listing failed:")
            print(result.stderr)
            sys.exit(1)
        sample = {"smb_listing": result.stdout}
        print("SMB share listing OK.")
    else:
        print("ERROR: source type must be sharepoint, smb or confluence.")
        sys.exit(1)

    with open(sample_output_path, "w", encoding="utf-8") as handle:
        json.dump(sample, handle, indent=2)
    print(f"Done. Sample saved to {sample_output_path} (token NOT included).")


if __name__ == "__main__":
    main()

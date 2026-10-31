#!/usr/bin/env python3
"""Send a change alert (email and/or Teams) for today's clause changes.

Tracker task: P4-10 (change alert)
Prompts: db_host, db_port, db_name, db_user, db_password (secret),
         alert_channel (email/teams/both), smtp_relay_host, smtp_relay_port,
         alert_from_address, alert_to_addresses, teams_webhook_url (secret),
         send_test_only (yes/no)
Prerequisites: clause_changes rows from 04-detect-standard-changes.py,
  psql on PATH. Uses only the standard library (smtplib, urllib).
  Webhook URLs and passwords are never printed or saved.
"""
import getpass
import json
import smtplib
import subprocess
import sys
import tempfile
from email.message import EmailMessage
import urllib.request


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_optional(label, default=""):
    return input(f"{label} [{default}]: ").strip() or default


def prompt_secret(label):
    while True:
        value = getpass.getpass(f"{label}: ").strip()
        if value:
            return value
        print("A value is required.")


def fetch_today_changes(host, port, db, user, password):
    sql = ("SELECT standard, chapter, clause_no, change_type, new_version "
           "FROM clause_changes WHERE detected_at >= CURRENT_DATE ORDER BY standard, clause_no;")
    with tempfile.NamedTemporaryFile("w", suffix=".sql", delete=False) as handle:
        handle.write(sql)
        sql_path = handle.name
    env = {"PGPASSWORD": password, "PATH": "/usr/bin:/bin:/usr/local/bin"}
    result = subprocess.run(
        ["psql", "-h", host, "-p", port, "-U", user, "-d", db,
         "-t", "-A", "-F", "|", "-f", sql_path],
        capture_output=True, text=True, env=env)
    if result.returncode != 0:
        print(result.stderr)
        sys.exit(1)
    changes = []
    for line in result.stdout.strip().splitlines():
        if not line.strip():
            continue
        standard, chapter, clause_no, change_type, new_version = line.split("|")
        changes.append(f"- {standard} {chapter} {clause_no}: {change_type} (version {new_version})")
    return changes


def main():
    print("=== Step 1: Alert settings ===")
    db_host = prompt("Database host", "PRD-AIRAG-DB01")
    db_port = prompt("Database port", "5432")
    db_name = prompt("Database name", "airag_standards")
    db_user = prompt("Database user", "airag_app")
    db_password = prompt_secret("Database password (not echoed, never saved)")
    alert_channel = prompt("Alert channel (email/teams/both)", "email").lower()
    smtp_relay_host = ""
    smtp_relay_port = "25"
    alert_from_address = ""
    alert_to_addresses = ""
    teams_webhook_url = ""
    if alert_channel in ("email", "both"):
        smtp_relay_host = prompt("SMTP relay host", "smtp.internal")
        smtp_relay_port = prompt("SMTP relay port", "25")
        alert_from_address = prompt("From address", "airag@internal")
        alert_to_addresses = prompt("To addresses, comma-separated", "qa@internal")
    if alert_channel in ("teams", "both"):
        teams_webhook_url = prompt_secret("Teams webhook URL (not echoed, never saved)")
    send_test_only = prompt("Send a test alert instead of real changes? (yes/no)", "no")

    print("=== Step 2: Building the message ===")
    if send_test_only == "yes":
        changes = ["- TEST: FDA eCFR Title 21 Part 211 211.100: modified (version 2026-01-01)"]
    else:
        changes = fetch_today_changes(db_host, db_port, db_name, db_user, db_password)
    if not changes:
        print("No clause changes today - no alert sent.")
        return
    body = ("AI-RAG standards change alert\n\n"
            "The following clauses changed. Review affected documents in the app:\n\n"
            + "\n".join(changes))

    if alert_channel in ("email", "both"):
        message = EmailMessage()
        message["Subject"] = "AI-RAG: standard clause changes detected"
        message["From"] = alert_from_address
        message["To"] = alert_to_addresses
        message.set_content(body)
        with smtplib.SMTP(smtp_relay_host, int(smtp_relay_port)) as smtp:
            smtp.send_message(message)
        print(f"Email alert sent to {alert_to_addresses}.")
    if alert_channel in ("teams", "both"):
        payload = json.dumps({"text": body}).encode("utf-8")
        request = urllib.request.Request(teams_webhook_url, data=payload,
                                         headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(request, timeout=30) as response:
            print(f"Teams alert sent (HTTP {response.status}).")
    print("Done.")


if __name__ == "__main__":
    main()

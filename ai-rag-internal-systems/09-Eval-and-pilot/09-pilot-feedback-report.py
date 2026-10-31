#!/usr/bin/env python3
"""Weekly pilot feedback report from the audit database.

Tracker task: P9-06 (run 4-week pilot, collect feedback weekly)
Prompts: db_host, db_port, db_name, db_user, db_password (secret),
         report_weeks, output_report_path
Prerequisites: audit tables from 07-create-audit-log.sql, psql on PATH,
  feedback rows from the app (08-app-main.py /feedback endpoint).
Output: a Markdown report with total feedback, % helpful, and a per-week
  breakdown. File the report with the weekly pilot review notes.
"""
import getpass
import subprocess
import sys
import tempfile


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_int(label, default):
    while True:
        raw = input(f"{label} [{default}]: ").strip() or str(default)
        if raw.isdigit() and int(raw) > 0:
            return int(raw)
        print("Enter a positive whole number.")


def prompt_secret(label):
    while True:
        value = getpass.getpass(f"{label}: ").strip()
        if value:
            return value
        print("A value is required.")


def run_psql(host, port, db, user, password, sql):
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
    return result.stdout.strip()


def main():
    print("=== Step 1: Report settings ===")
    db_host = prompt("Database host", "PRD-AIRAG-DB01")
    db_port = prompt("Database port", "5432")
    db_name = prompt("Database name", "airag_audit")
    db_user = prompt("Database user", "airag_app")
    db_password = prompt_secret("Database password (not echoed, never saved)")
    report_weeks = prompt_int("How many past weeks to include", 4)
    output_report_path = prompt("Report output path", "/tmp/airag-pilot-feedback.md")

    print("=== Step 2: Querying feedback ===")
    totals = run_psql(
        db_host, db_port, db_name, db_user, db_password,
        "SELECT COUNT(*), SUM(CASE WHEN helpful THEN 1 ELSE 0 END) FROM feedback;")
    weekly = run_psql(
        db_host, db_port, db_name, db_user, db_password,
        f"SELECT date_trunc('week', created_at)::date, COUNT(*), "
        f"SUM(CASE WHEN helpful THEN 1 ELSE 0 END) FROM feedback "
        f"WHERE created_at >= now() - INTERVAL '{report_weeks} weeks' "
        f"GROUP BY 1 ORDER BY 1;")
    comments = run_psql(
        db_host, db_port, db_name, db_user, db_password,
        f"SELECT comment FROM feedback WHERE comment IS NOT NULL AND comment <> '' "
        f"AND created_at >= now() - INTERVAL '{report_weeks} weeks' ORDER BY created_at DESC LIMIT 20;")

    total_count, helpful_count = (totals.split("|") + ["0", "0"])[:2]
    total_count = int(total_count or 0)
    helpful_count = int(helpful_count or 0)
    report = (
        "# AI-RAG pilot feedback report\n\n"
        f"- Feedback rows: {total_count}\n"
        f"- Helpful: {helpful_count} ({100 * helpful_count // total_count if total_count else 0}%)\n\n"
        "## Per week (week, total, helpful)\n\n"
        + "".join(f"- {line}\n" for line in weekly.splitlines()) + "\n"
        "## Recent comments\n\n"
        + "".join(f"- {line}\n" for line in comments.splitlines())
    )
    with open(output_report_path, "w", encoding="utf-8") as handle:
        handle.write(report)
    print(report)
    print(f"Report saved to {output_report_path}.")
    print("Triage the comments into tickets at the weekly pilot review.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Detect clause changes: compare freshly fetched clauses to the registry.

Tracker task: P4-09 (change detector)
Prompts: new_clauses_jsonl_path, db_host, db_port, db_name, db_user,
         db_password (secret), alert_on_change (yes/no)
Prerequisites: clause registry created; a fresh clause JSONL from the
  phase-04 fetch/parse scripts. Uses one psql run per comparison step.
How it works: loads the new clauses into a temp staging table, then
  compares against the latest version per (standard, chapter, clause_no):
  new key -> 'added'; missing key -> 'removed'; different text -> 'modified'.
  Every change is written to clause_changes (append-only).
"""
import getpass
import json
import os
import subprocess
import sys
import tempfile


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


def run_psql_file(host, port, db, user, password, sql_path):
    env = {"PGPASSWORD": password, "PATH": "/usr/bin:/bin:/usr/local/bin"}
    result = subprocess.run(
        ["psql", "-h", host, "-p", port, "-U", user, "-d", db,
         "-v", "ON_ERROR_STOP=1", "-f", sql_path],
        capture_output=True, text=True, env=env)
    if result.returncode != 0:
        print(result.stderr)
        sys.exit(1)
    return result.stdout


def main():
    print("=== Step 1: Diff settings ===")
    new_clauses_jsonl_path = prompt("Fresh clauses JSONL (example — replace with your value)", "/data/staging/clauses-ecfr.jsonl")
    db_host = prompt("Database host (example — replace with your value)", "PRD-AIRAG-DB01")
    db_port = prompt("Database port (example — replace with your value)", "5432")
    db_name = prompt("Database name (example — replace with your value)", "airag_standards")
    db_user = prompt("Database user (example — replace with your value)", "airag_app")
    db_password = prompt_secret("Database password (not echoed, never saved)")
    alert_on_change = prompt("Print a change summary for alerting? (yes/no) (example — replace with your value)", "yes")

    print("=== Step 2: Comparing ===")
    rows = []
    with open(new_clauses_jsonl_path, encoding="utf-8") as handle:
        for line in handle:
            record = json.loads(line)
            rows.append((record["standard"], record["chapter"], record["clause_no"],
                         record["text"].replace("\t", " ").replace("\n", " "),
                         record["version"], record.get("source_url", "")))
    with tempfile.TemporaryDirectory() as tmp_dir:
        tsv_path = os.path.join(tmp_dir, "new_clauses.tsv")
        with open(tsv_path, "w", encoding="utf-8") as tmp:
            for row in rows:
                tmp.write("\t".join(row) + "\n")
        sql_path = os.path.join(tmp_dir, "diff.sql")
        with open(sql_path, "w", encoding="utf-8") as sql:
            sql.write(
                "CREATE TEMP TABLE new_clauses (standard TEXT, chapter TEXT, clause_no TEXT, "
                "text TEXT, version DATE, source_url TEXT);\n"
                f"\\COPY new_clauses FROM '{tsv_path}' WITH (FORMAT text)\n"
                # Latest registered version per clause key.
                "CREATE TEMP TABLE latest_clauses AS "
                "SELECT DISTINCT ON (standard, chapter, clause_no) standard, chapter, clause_no, text, version "
                "FROM clauses ORDER BY standard, chapter, clause_no, version DESC;\n"
                # Added + modified.
                "INSERT INTO clause_changes (standard, chapter, clause_no, old_version, new_version, change_type, old_text, new_text) "
                "SELECT n.standard, n.chapter, n.clause_no, l.version, n.version, "
                "CASE WHEN l.clause_no IS NULL THEN 'added' ELSE 'modified' END, l.text, n.text "
                "FROM new_clauses n LEFT JOIN latest_clauses l "
                "ON n.standard = l.standard AND n.chapter = l.chapter AND n.clause_no = l.clause_no "
                "WHERE l.clause_no IS NULL OR l.text IS DISTINCT FROM n.text;\n"
                # Removed (present in registry, absent from the fresh fetch).
                "INSERT INTO clause_changes (standard, chapter, clause_no, old_version, new_version, change_type, old_text, new_text) "
                "SELECT l.standard, l.chapter, l.clause_no, l.version, CURRENT_DATE, 'removed', l.text, NULL "
                "FROM latest_clauses l LEFT JOIN new_clauses n "
                "ON n.standard = l.standard AND n.chapter = l.chapter AND n.clause_no = l.clause_no "
                "WHERE n.clause_no IS NULL;\n"
                # Also store the fresh clauses as the new version rows.
                "INSERT INTO clauses (standard, chapter, clause_no, text, version, source_url) "
                "SELECT standard, chapter, clause_no, text, version, source_url FROM new_clauses "
                "ON CONFLICT (standard, chapter, clause_no, version) DO NOTHING;\n"
                "SELECT change_type, COUNT(*) FROM clause_changes "
                "WHERE detected_at >= CURRENT_DATE GROUP BY change_type;\n")
        summary = run_psql_file(db_host, db_port, db_name, db_user, db_password, sql_path)
    print("=== Step 3: Change summary (today) ===")
    print(summary)
    if alert_on_change == "yes":
        print("If any rows above show changes, run 04-send-change-alert.py to notify QA.")
    print("Done.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Build the AD-group -> document-prefix access map interactively.

Tracker task: P7-02 (map AD groups to document access sets)
Prompts: db_host, db_port, db_name, db_user, db_password (secret),
         access_map_path
Prerequisites: access_map table - created by this script if missing.
  Uses psql on PATH (no Python DB driver needed).
How it works: the operator adds groups one at a time, each with the
  document ID prefixes that group may see (e.g. SOP-QA, SOP-PROD).
  The map is saved as JSON (for 05-permission-filter.py) AND written to
  the airag_docs.access_map table. QA must review and sign off the map.
"""
import getpass
import json
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
    print("=== Step 1: Database settings ===")
    db_host = prompt("Database host", "PRD-AIRAG-DB01")
    db_port = prompt("Database port", "5432")
    db_name = prompt("Database name", "airag_docs")
    db_user = prompt("Database user", "airag_app")
    db_password = prompt_secret("Database password (not echoed, never saved)")
    access_map_path = prompt("Access map JSON path", "/opt/airag/access-map.json")

    print("=== Step 2: Build the map (blank group name = finished) ===")
    access_map = {}
    while True:
        group = input("AD group name (blank to finish): ").strip()
        if not group:
            break
        prefixes = input(f"  Doc ID prefixes for {group}, comma-separated: ").strip()
        access_map[group] = [entry.strip() for entry in prefixes.split(",") if entry.strip()]
    if not access_map:
        print("No groups entered - nothing to save.")
        return
    with open(access_map_path, "w", encoding="utf-8") as handle:
        json.dump(access_map, handle, indent=2)
    print(f"Map saved to {access_map_path}.")

    print("=== Step 3: Write the map to the database ===")
    with tempfile.NamedTemporaryFile("w", suffix=".sql", delete=False) as sql:
        sql.write("CREATE TABLE IF NOT EXISTS access_map "
                  "(ad_group TEXT PRIMARY KEY, doc_prefixes TEXT[] NOT NULL);\n")
        for group, prefixes in access_map.items():
            escaped = "{" + ",".join(f'"{prefix}"' for prefix in prefixes) + "}"
            safe_group = group.replace("'", "''")
            sql.write(f"INSERT INTO access_map VALUES ('{safe_group}', '{escaped}') "
                      f"ON CONFLICT (ad_group) DO UPDATE SET doc_prefixes = EXCLUDED.doc_prefixes;\n")
        sql.write("SELECT ad_group, doc_prefixes FROM access_map ORDER BY ad_group;\n")
        sql_path = sql.name
    print(run_psql_file(db_host, db_port, db_name, db_user, db_password, sql_path))
    print("Done. QA must review and sign off this map (P7-02 acceptance).")
    print("Test with users from 2 different groups before go-live.")


if __name__ == "__main__":
    main()

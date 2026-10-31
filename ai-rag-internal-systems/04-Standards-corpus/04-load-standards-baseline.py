#!/usr/bin/env python3
"""Load clause JSONL files into the clause registry (baseline load).

Tracker task: P4-08 (load baseline versions of all three families)
Prompts: clauses_jsonl_paths (comma-separated), db_host, db_port, db_name,
         db_user, db_password (secret)
Prerequisites: clause registry created (04-create-clause-registry.sql),
  psql on PATH. Uses one psql run per file (temp staging table + COPY +
  upsert in a single session) - no Python DB driver needed.
  Re-running is safe: rows are skipped on (standard, chapter, clause_no,
  version) conflicts.
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
    print("=== Step 1: Load settings ===")
    clauses_jsonl_paths = prompt("Clause JSONL files, comma-separated",
                                 "/data/staging/clauses-ecfr.jsonl,/data/staging/clauses-standards.jsonl")
    db_host = prompt("Database host", "PRD-AIRAG-DB01")
    db_port = prompt("Database port", "5432")
    db_name = prompt("Database name", "airag_standards")
    db_user = prompt("Database user", "airag_app")
    db_password = prompt_secret("Database password (not echoed, never saved)")

    print("=== Step 2: Loading ===")
    total = 0
    for path in [entry.strip() for entry in clauses_jsonl_paths.split(",") if entry.strip()]:
        rows = []
        with open(path, encoding="utf-8") as handle:
            for line in handle:
                record = json.loads(line)
                rows.append((record["standard"], record["chapter"], record["clause_no"],
                             record["text"].replace("\t", " ").replace("\n", " "),
                             record["version"], record.get("source_url", "")))
        with tempfile.TemporaryDirectory() as tmp_dir:
            tsv_path = os.path.join(tmp_dir, "clauses.tsv")
            with open(tsv_path, "w", encoding="utf-8") as tmp:
                for row in rows:
                    tmp.write("\t".join(row) + "\n")
            sql_path = os.path.join(tmp_dir, "load.sql")
            with open(sql_path, "w", encoding="utf-8") as sql:
                sql.write(
                    "CREATE TEMP TABLE clauses_stage (standard TEXT, chapter TEXT, "
                    "clause_no TEXT, text TEXT, version DATE, source_url TEXT);\n"
                    f"\\COPY clauses_stage FROM '{tsv_path}' WITH (FORMAT text)\n"
                    "INSERT INTO clauses (standard, chapter, clause_no, text, version, source_url) "
                    "SELECT standard, chapter, clause_no, text, version, source_url FROM clauses_stage "
                    "ON CONFLICT (standard, chapter, clause_no, version) DO NOTHING;\n")
            run_psql_file(db_host, db_port, db_name, db_user, db_password, sql_path)
        print(f"  {path}: {len(rows)} rows staged")
        total += len(rows)

    print("=== Step 3: Verify ===")
    with tempfile.NamedTemporaryFile("w", suffix=".sql", delete=False) as sql:
        sql.write("SELECT standard, COUNT(*) FROM clauses GROUP BY standard ORDER BY standard;\n")
        sql_path = sql.name
    print(run_psql_file(db_host, db_port, db_name, db_user, db_password, sql_path))
    os.unlink(sql_path)
    print(f"Done. {total} clause rows processed (baseline).")


if __name__ == "__main__":
    main()

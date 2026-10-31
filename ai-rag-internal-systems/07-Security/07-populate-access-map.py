#!/usr/bin/env python3
"""Build the AD-group -> document-prefix access map interactively.

Tracker task: P7-02 (map AD groups to document access sets)
Prompts: db_host, db_port, db_name, db_user, db_password (secret),
         access_map_path, then in a loop: group, prefixes
Prerequisites: access_map table - created by this script if missing.
  Requires psycopg (psycopg 3, or psycopg2 as a fallback). All inserts
  use parameterized queries (%s placeholders) with the group name and
  the prefix list passed as parameters - never interpolated into SQL.
  Passing the list as a parameter lets the driver encode the TEXT[]
  value, so a '"' or '\\' in a prefix cannot break or inject into an
  array literal. No temp SQL file is created at all (the previous
  version's temp file, which held interpolated values, is gone).
  The DB password reaches libpq via the PGPASSWORD environment variable
  only: it is never passed on a command line, never printed, never
  logged, and never saved. It is removed from the environment once the
  database work is done.
How it works: the operator adds groups one at a time, each with the
  document ID prefixes that group may see (e.g. SOP-QA, SOP-PROD).
  The map is saved as JSON (for 05-permission-filter.py) AND written to
  the airag_docs.access_map table. QA must review and sign off the map.
Example defaults are labelled "(example)" in their prompts; confirm
  every default at runtime.
"""
import getpass
import json
import os
import sys

try:
    import psycopg
except ImportError:
    try:
        import psycopg2 as psycopg
    except ImportError:
        psycopg = None


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


def write_access_map(host, port, db, user, password, access_map):
    """Write the map with parameterized queries; return the stored rows."""
    if psycopg is None:
        print("ERROR: psycopg is not installed (pip install psycopg).")
        sys.exit(1)
    # PGPASSWORD env only: libpq reads it at connect time, so no
    # password argument is passed to connect() and none appears in argv.
    os.environ["PGPASSWORD"] = password
    try:
        conn = psycopg.connect(host=host, port=port, dbname=db, user=user)
        try:
            with conn.cursor() as cur:
                cur.execute(
                    "CREATE TABLE IF NOT EXISTS access_map "
                    "(ad_group TEXT PRIMARY KEY, doc_prefixes TEXT[] NOT NULL)")
                for group, prefixes in access_map.items():
                    cur.execute(
                        "INSERT INTO access_map (ad_group, doc_prefixes) "
                        "VALUES (%s, %s) "
                        "ON CONFLICT (ad_group) DO UPDATE SET "
                        "doc_prefixes = EXCLUDED.doc_prefixes",
                        (group, prefixes))
                cur.execute(
                    "SELECT ad_group, doc_prefixes FROM access_map "
                    "ORDER BY ad_group")
                rows = cur.fetchall()
            conn.commit()
            return rows
        finally:
            conn.close()
    finally:
        os.environ.pop("PGPASSWORD", None)


def main():
    print("=== Step 1: Database settings ===")
    db_host = prompt("Database host (example)", "PRD-AIRAG-DB01")
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
    rows = write_access_map(db_host, db_port, db_name, db_user, db_password, access_map)
    for ad_group, doc_prefixes in rows:
        print(f"  {ad_group}: {doc_prefixes}")
    print("Done. QA must review and sign off this map (P7-02 acceptance).")
    print("Test with users from 2 different groups before go-live.")


if __name__ == "__main__":
    main()

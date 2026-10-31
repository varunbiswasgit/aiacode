#!/usr/bin/env python3
"""Build the chunks_fts Postgres table from chunk JSONL.

Tracker task: P3-13 (build chunks_fts for hybrid retrieval)
Prompts: chunks_jsonl_path, db_host, db_port, db_name, db_user,
         db_password (secret), load_mode (upsert/rebuild)
Prerequisites: Python 3.12, psql on PATH, a Postgres database that
  05-hybrid-retrieve.py will query. Input JSONL is the chunk format
  produced by 03-ingest-corpus.py: one JSON object per line with
  {corpus, doc_id, source_path, chunk_index, text}.
Logic: read and validate the JSONL (rows missing doc_id, chunk_index
  or text are skipped and counted, never loaded); stage the rows with
  a CSV file loaded by psql COPY into a temporary staging table; then
  INSERT ... SELECT into chunks_fts, computing the tsvector on the
  server with to_tsvector('english', text). No chunk text is ever
  interpolated into SQL text: values travel only through the COPY
  file, and the COPY file path is passed to psql as a -v variable
  (:'csv_path'), never spliced into a SQL string by this script.
  Re-running in upsert mode is safe: rows are matched on
  (doc_id, chunk_index) and updated in place. Rebuild mode drops
  chunks_fts and recreates it from the JSONL only, after an explicit
  confirmation prompt. This script accepts no command-line arguments:
  the password is never taken from argv.
Configuration (defaults shown are examples only):
  - chunks_jsonl_path: /data/staging/chunks-pilot.jsonl
    (example — replace with your value)
  - db_host: localhost (example — replace with your value)
  - db_port: 5432 (example — replace with your value)
  - db_name: airag_docs (example — replace with your value)
  - db_user: airag_app (example — replace with your value)
  - load_mode: upsert (example — replace with your value)
  The tsvector configuration ('english') is fixed, not prompted: it
  must match the plainto_tsquery('english', ...) calls in
  05-hybrid-retrieve.py.
Dependencies: Python 3 standard library only; psql on PATH; no
  Python database driver is needed. The database password is taken
  from the PGPASSWORD environment variable when set, otherwise read
  with getpass (not echoed, never saved, never printed or logged);
  it is passed to psql only through the PGPASSWORD environment
  variable of the psql subprocess, never as an argument.
Output: the chunks_fts table plus a GIN index on tsv, and a row count
  printed on completion. The table is the one read by
  05-RAG-core/05-hybrid-retrieve.py.
Schema (chunks_fts):
  doc_id      TEXT NOT NULL,
  chunk_index INTEGER NOT NULL,
  corpus      TEXT,
  source_path TEXT,
  text        TEXT NOT NULL,
  tsv         tsvector NOT NULL,
  PRIMARY KEY (doc_id, chunk_index)
  plus GIN index chunks_fts_tsv_idx on (tsv).
  This matches 05-hybrid-retrieve.py exactly: it selects
  doc_id, chunk_index, text from chunks_fts where
  tsv @@ plainto_tsquery('english', <query>) ordered by
  ts_rank(tsv, plainto_tsquery('english', <query>)).
  corpus and source_path are carried over from the JSONL for
  provenance; the retrieval query ignores them.
"""
import csv
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


def prompt_int(label, default):
    while True:
        raw = input(f"{label} [{default}]: ").strip() or str(default)
        if raw.isdigit() and int(raw) > 0:
            return int(raw)
        print("Enter a positive whole number.")


def prompt_password():
    """Password from PGPASSWORD env, else getpass. Never argv, never logged."""
    env_password = os.environ.get("PGPASSWORD", "").strip()
    if env_password:
        print("Using Postgres password from the PGPASSWORD environment variable.")
        return env_password
    while True:
        value = getpass.getpass("Postgres password (not echoed, never saved): ").strip()
        if value:
            return value
        print("A value is required.")


def prompt_mode():
    while True:
        value = prompt("Load mode (upsert/rebuild)", "upsert").lower()
        if value in ("upsert", "rebuild"):
            return value
        print("Mode must be upsert or rebuild.")


def prompt_confirm(label):
    while True:
        value = input(f"{label} [y/N]: ").strip().lower()
        if value in ("y", "yes"):
            return True
        if value in ("n", "no", ""):
            return False
        print("Enter y or n.")


def run_psql(host, port, db, user, password, args):
    """Run psql with the password only in the subprocess environment."""
    env = {"PGPASSWORD": password, "PATH": "/usr/bin:/bin:/usr/local/bin"}
    result = subprocess.run(
        ["psql", "-h", host, "-p", str(port), "-U", user, "-d", db,
         "-v", "ON_ERROR_STOP=1"] + args,
        capture_output=True, text=True, env=env)
    if result.returncode != 0:
        print(result.stderr)
        sys.exit(1)
    return result.stdout


def load_jsonl(path):
    """Return (rows, skipped). Rows are (doc_id, chunk_index, corpus, source_path, text)."""
    rows = []
    skipped = 0
    with open(path, encoding="utf-8") as handle:
        for line_number, line in enumerate(handle, 1):
            if not line.strip():
                continue
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                print(f"  WARNING: line {line_number}: not valid JSON, skipped.")
                skipped += 1
                continue
            doc_id = record.get("doc_id")
            chunk_index = record.get("chunk_index")
            text = record.get("text")
            if (not isinstance(doc_id, str) or not doc_id.strip()
                    or not isinstance(chunk_index, int) or isinstance(chunk_index, bool)
                    or not isinstance(text, str) or not text.strip()):
                print(f"  WARNING: line {line_number}: missing doc_id, chunk_index "
                      "or text, skipped.")
                skipped += 1
                continue
            rows.append((doc_id, chunk_index, record.get("corpus") or "",
                         record.get("source_path") or "", text))
    return rows, skipped


def main():
    if len(sys.argv) > 1:
        print("ERROR: this script takes no command-line arguments; "
              "all values are prompted at runtime.")
        raise SystemExit(1)

    print("=== Step 1: Build settings ===")
    chunks_jsonl_path = prompt(
        "Chunks JSONL (example — replace with your value)",
        "/data/staging/chunks-pilot.jsonl")
    db_host = prompt("Postgres host (example — replace with your value)", "localhost")
    db_port = prompt_int("Postgres port (example — replace with your value)", 5432)
    db_name = prompt("Postgres database (example — replace with your value)", "airag_docs")
    db_user = prompt("Postgres user (example — replace with your value)", "airag_app")
    db_password = prompt_password()
    load_mode = prompt_mode()
    if load_mode == "rebuild":
        print("WARNING: rebuild drops the existing chunks_fts table and "
              "recreates it from the JSONL only.")
        if not prompt_confirm("Drop chunks_fts and rebuild it now?"):
            print("Aborted; chunks_fts left unchanged.")
            raise SystemExit(1)

    print(f"=== Step 2: Reading {chunks_jsonl_path} ===")
    if not os.path.exists(chunks_jsonl_path):
        print(f"ERROR: JSONL file not found: {chunks_jsonl_path}")
        raise SystemExit(1)
    rows, skipped = load_jsonl(chunks_jsonl_path)
    print(f"  {len(rows)} valid chunks, {skipped} lines skipped.")
    if not rows:
        print("ERROR: no valid chunks to load.")
        raise SystemExit(1)

    print(f"=== Step 3: Loading chunks_fts ({load_mode}) ===")
    if load_mode == "rebuild":
        ddl = ("DROP TABLE IF EXISTS chunks_fts;\n"
               "CREATE TABLE chunks_fts (\n"
               "  doc_id TEXT NOT NULL,\n"
               "  chunk_index INTEGER NOT NULL,\n"
               "  corpus TEXT,\n"
               "  source_path TEXT,\n"
               "  text TEXT NOT NULL,\n"
               "  tsv tsvector NOT NULL,\n"
               "  PRIMARY KEY (doc_id, chunk_index)\n"
               ");\n"
               "CREATE INDEX chunks_fts_tsv_idx ON chunks_fts USING GIN (tsv);\n")
    else:
        ddl = ("CREATE TABLE IF NOT EXISTS chunks_fts (\n"
               "  doc_id TEXT NOT NULL,\n"
               "  chunk_index INTEGER NOT NULL,\n"
               "  corpus TEXT,\n"
               "  source_path TEXT,\n"
               "  text TEXT NOT NULL,\n"
               "  tsv tsvector NOT NULL,\n"
               "  PRIMARY KEY (doc_id, chunk_index)\n"
               ");\n"
               "CREATE INDEX IF NOT EXISTS chunks_fts_tsv_idx "
               "ON chunks_fts USING GIN (tsv);\n"
               "-- NOTE: if chunks_fts was created by the old ad-hoc\n"
               "-- CREATE TABLE AS statement (no primary key), run this\n"
               "-- script once in rebuild mode to add the primary key\n"
               "-- and the GIN index; upsert works from then on.\n")
    # Fixed SQL text: no chunk values are interpolated into it. The
    # staging data arrives via COPY from the CSV file whose path is
    # handed to psql as a -v variable (:'csv_path').
    sql = (ddl
           + "CREATE TEMP TABLE chunks_stage (\n"
           + "  doc_id TEXT NOT NULL,\n"
           + "  chunk_index INTEGER NOT NULL,\n"
           + "  corpus TEXT,\n"
           + "  source_path TEXT,\n"
           + "  text TEXT NOT NULL\n"
           + ");\n"
           + "\\COPY chunks_stage FROM :'csv_path' WITH (FORMAT csv, HEADER true)\n"
           + "INSERT INTO chunks_fts (doc_id, chunk_index, corpus, source_path, text, tsv)\n"
           + "SELECT doc_id, chunk_index, corpus, source_path, text,\n"
           + "       to_tsvector('english', text)\n"
           + "FROM chunks_stage\n"
           + "ON CONFLICT (doc_id, chunk_index) DO UPDATE SET\n"
           + "  corpus = EXCLUDED.corpus,\n"
           + "  source_path = EXCLUDED.source_path,\n"
           + "  text = EXCLUDED.text,\n"
           + "  tsv = EXCLUDED.tsv;\n"
           + "DROP TABLE chunks_stage;\n")
    with tempfile.TemporaryDirectory() as tmp_dir:
        csv_path = os.path.join(tmp_dir, "chunks.csv")
        with open(csv_path, "w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(["doc_id", "chunk_index", "corpus", "source_path", "text"])
            writer.writerows(rows)
        sql_path = os.path.join(tmp_dir, "load.sql")
        with open(sql_path, "w", encoding="utf-8") as handle:
            handle.write(sql)
        run_psql(db_host, db_port, db_name, db_user, db_password,
                 ["-v", f"csv_path={csv_path}", "-f", sql_path])

    print("=== Step 4: Verify ===")
    count = run_psql(db_host, db_port, db_name, db_user, db_password,
                     ["-t", "-A", "-c", "SELECT COUNT(*) FROM chunks_fts;"]).strip()
    print(f"Done. chunks_fts now holds {count} rows (mode: {load_mode}).")
    print("Next: 05-hybrid-retrieve.py reads chunks_fts for keyword search.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Send a change alert (email and/or Teams) for today's clause changes.

Tracker task: P4-10 (change alert)
Prompts: db_host, db_port, db_name, db_user, db_password (secret),
         mappings_jsonl_path, suggestions_jsonl_path (optional),
         alert_channel (email/teams/both), smtp_relay_host, smtp_relay_port,
         alert_from_address, alert_to_addresses, teams_webhook_url (secret),
         send_test_only (yes/no)
Prerequisites: clause_changes rows from 04-detect-standard-changes.py,
  psql on PATH, and mappings JSONL from 03-map-documents-to-clauses.py
  (rows: doc_id, chunk_index, clause_key, status). The alert joins each
  change to that file and lists the actual affected documents, so
  recipients are never sent to a view that does not exist: the alert
  itself carries the document list, and review happens in the app's
  existing review queue. No link to any other view is included.
  An optional suggestions JSONL (rows: doc_id, clause_key and/or
  impact_rating, where impact_rating is High/Medium/Low from gap
  analysis) adds an impact rating per document when one is available;
  when no rating is on file the alert says so instead of inventing one.
  This script suggests only: it never edits a controlled document.
  Uses only the standard library (smtplib, urllib).
  The database password reaches psql through the PGPASSWORD
  environment variable only, never as a command-line argument, and
  passwords and webhook URLs are never printed, logged, or saved.
  The temporary SQL file is always deleted (try/finally).
"""
import getpass
import json
import os
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
    sql_path = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".sql", delete=False) as handle:
            handle.write(sql)
            sql_path = handle.name
        # Password via the PGPASSWORD environment variable only; it is
        # never passed as a psql argument and never printed.
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
            changes.append({"standard": standard, "chapter": chapter,
                            "clause_no": clause_no, "change_type": change_type,
                            "new_version": new_version})
        return changes
    finally:
        if sql_path:
            try:
                os.unlink(sql_path)
            except FileNotFoundError:
                pass


def load_jsonl(path, description):
    """Load a JSONL file, skipping malformed lines with a warning."""
    rows = []
    try:
        with open(path, encoding="utf-8") as handle:
            for number, line in enumerate(handle, 1):
                if not line.strip():
                    continue
                try:
                    rows.append(json.loads(line))
                except json.JSONDecodeError:
                    print(f"WARNING: skipping malformed {description} line {number}.")
    except FileNotFoundError:
        print(f"WARNING: {description} file not found: {path}. "
              "Affected documents / impact ratings from it will be unavailable.")
    return rows


def clause_matches(change, clause_key):
    """True when a mappings clause_key refers to this change's clause.

    mappings.jsonl clause_keys take two shapes (see
    03-map-documents-to-clauses.py): the keyword form
    "<standard>:<clause_no>" (e.g. "FDA eCFR:211.100") and the citation
    form "<standard>:<citation text>" (e.g. "FDA eCFR:21 CFR 211.100",
    "ICH:ICH Q7 5.2"). The standard must match, and the clause number
    must be the whole key suffix or its final token, so "211.10" never
    matches "211.100". The leaf clause_no (see 04-parse-ecfr-clauses.py)
    is what is compared, never a full ancestor path.
    """
    if ":" not in clause_key:
        return False
    standard, suffix = clause_key.split(":", 1)
    if standard.strip() != change["standard"]:
        return False
    suffix = suffix.strip()
    if not suffix:
        return False
    clause_no = change["clause_no"]
    return suffix == clause_no or suffix.split()[-1] == clause_no


def load_impacts(path):
    """Build impact lookups from an optional suggestions JSONL.

    Returns (by_doc_clause, by_doc): by_doc_clause maps
    (doc_id, clause_key) to a High/Medium/Low rating, by_doc maps a
    doc_id to its rating when no clause-specific rating is on file.
    Rows with any other rating value are ignored, never relabelled.
    An empty path means no suggestions file: both lookups stay empty.
    """
    by_doc_clause = {}
    by_doc = {}
    if not path:
        return by_doc_clause, by_doc
    for row in load_jsonl(path, "suggestions"):
        rating = str(row.get("impact_rating", "")).strip()
        if rating.lower() not in ("high", "medium", "low"):
            continue
        rating = rating.capitalize()
        doc_id = row.get("doc_id")
        if not doc_id:
            continue
        clause_key = row.get("clause_key")
        if clause_key:
            by_doc_clause[(doc_id, clause_key)] = rating
        else:
            by_doc[doc_id] = rating
    return by_doc_clause, by_doc


def impact_for(doc_id, mapping, change, by_doc_clause, by_doc):
    """Best available impact rating for one document/clause pair."""
    clause_key = mapping.get("clause_key", "")
    if (doc_id, clause_key) in by_doc_clause:
        return by_doc_clause[(doc_id, clause_key)]
    if (doc_id, change["clause_no"]) in by_doc_clause:
        return by_doc_clause[(doc_id, change["clause_no"])]
    return by_doc.get(doc_id, "not assessed")


def build_alert(changes, mappings, by_doc_clause, by_doc, mappings_available):
    """Render the alert body: each change plus its affected documents."""
    lines = ["AI-RAG standards change alert",
             "",
             "The following clauses changed. Review the affected documents "
             "listed below in the app's review queue. This alert suggests "
             "only; it does not edit any controlled document.",
             ""]
    if not mappings_available:
        lines.append("NOTE: the mappings file was unavailable, so affected "
                     "documents could not be listed. Check the mappings "
                     "before acting on these changes.")
        lines.append("")
    for change in changes:
        lines.append(f"- {change['standard']} {change['chapter']} "
                     f"{change['clause_no']}: {change['change_type']} "
                     f"(version {change['new_version']})")
        affected = {}
        for mapping in mappings:
            clause_key = mapping.get("clause_key", "")
            doc_id = mapping.get("doc_id")
            if not doc_id or not clause_matches(change, clause_key):
                continue
            entry = affected.setdefault(doc_id, {"statuses": set(), "chunks": set()})
            if mapping.get("status"):
                entry["statuses"].add(mapping["status"])
            if mapping.get("chunk_index") is not None:
                entry["chunks"].add(str(mapping["chunk_index"]))
        if not affected:
            lines.append("  Affected documents: none mapped to this clause "
                         "in the mappings file.")
        else:
            lines.append("  Affected documents:")
            for doc_id in sorted(affected):
                entry = affected[doc_id]
                detail = []
                if entry["statuses"]:
                    detail.append(f"mapping: {', '.join(sorted(entry['statuses']))}")
                if entry["chunks"]:
                    detail.append(f"chunks: {', '.join(sorted(entry['chunks']))}")
                # Prefer a clause-specific rating when one is on file;
                # otherwise fall back to the document-level rating.
                rating = by_doc.get(doc_id, "not assessed")
                for mapping in mappings:
                    if (mapping.get("doc_id") == doc_id
                            and clause_matches(change, mapping.get("clause_key", ""))):
                        specific = impact_for(doc_id, mapping, change,
                                              by_doc_clause, by_doc)
                        if specific != "not assessed":
                            rating = specific
                            break
                detail.append(f"impact: {rating}")
                lines.append(f"  - {doc_id} ({'; '.join(detail)})")
    return "\n".join(lines)


def main():
    print("=== Step 1: Alert settings ===")
    db_host = prompt("Database host", "PRD-AIRAG-DB01")
    db_port = prompt("Database port", "5432")
    db_name = prompt("Database name", "airag_standards")
    db_user = prompt("Database user", "airag_app")
    db_password = prompt_secret("Database password (not echoed, never saved)")
    mappings_jsonl_path = prompt("Mappings JSONL", "/data/staging/mappings.jsonl")
    suggestions_jsonl_path = prompt_optional(
        "Suggestions JSONL with impact ratings (blank if none)", "")
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
        changes = [{"standard": "FDA eCFR", "chapter": "Title 21 Part 211",
                    "clause_no": "211.100", "change_type": "modified",
                    "new_version": "2026-01-01"}]
    else:
        changes = fetch_today_changes(db_host, db_port, db_name, db_user, db_password)
    if not changes:
        print("No clause changes today - no alert sent.")
        return
    mappings = load_jsonl(mappings_jsonl_path, "mappings")
    mappings_available = os.path.exists(mappings_jsonl_path)
    by_doc_clause, by_doc = load_impacts(suggestions_jsonl_path)
    body = build_alert(changes, mappings, by_doc_clause, by_doc, mappings_available)

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

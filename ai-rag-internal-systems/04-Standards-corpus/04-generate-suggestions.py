#!/usr/bin/env python3
"""Generate reviewer suggestions for internal documents affected by clause changes.

Tracker task: P4-12 (suggestion-generating step)

PURPOSE
-------
For each new/pending row in clause_changes, find the internal documents
mapped to that clause in mappings.jsonl, run the gap.txt and suggest.txt
prompt templates against a LOCAL model, and store one suggestion row per
affected document in the suggestions table for human review.

SUGGEST-NEVER-EDIT (HARD RULE)
-------------------------------
This script NEVER writes to, edits, or otherwise modifies any controlled
document. It writes ONLY to the suggestions store (the suggestions table
in the standards database) via INSERT. No UPDATE or DELETE is issued
against any table, and no document file is opened for writing. Every
suggestion it produces has status 'pending' and must be approved by a
human reviewer before anything changes in a controlled document.

PROMPTED VARIABLES (all values are asked at runtime; nothing is read
from command-line arguments, and no value below is a live credential)
--------------------------------------------------------------------
- db_host: standards database host.
- db_port: standards database port.
- db_name: standards database name.
- db_user: standards database user.
- db_password: standards database password. Read ONLY from the
  PGPASSWORD environment variable when it is set, otherwise asked with
  getpass (not echoed). It is NEVER taken from argv, NEVER printed,
  NEVER logged, and NEVER saved to a file.
- mappings_path: path to mappings.jsonl (the doc-to-clause map produced
  by 03-map-documents-to-clauses.py: rows of doc_id, chunk_index,
  clause_key, status).
- chunks_path (optional, blank = none): path to an internal chunks
  JSONL (rows of doc_id, chunk_index, text) used to supply the actual
  document-section text for the gap prompt. When blank, the gap prompt
  carries only the document/chunk identifiers plus the clause texts,
  and the model is told the section text was not supplied.
- templates_directory: directory holding gap.txt and suggest.txt
  (created by 05-create-prompt-templates.py).
- model_endpoint: full URL of the LOCAL model endpoint to POST to.
- model_name: model name sent to the local model endpoint.
- model_api_style: payload style of the local model endpoint:
  'ollama' or 'openai' (OpenAI-compatible chat completions).
- model_timeout_seconds: HTTP timeout in seconds for each model call.
- detected_since (optional, blank = no window): only consider
  clause_changes detected on or after this date (YYYY-MM-DD).
- batch_limit: maximum number of clause_changes to process in this run.
- include_needs_review_mappings (yes/no): whether mappings with status
  'needs_review' (keyword-only matches) count as affected documents.
  'confirmed' (citation) mappings always count.

LOGIC FLOW
----------
1. Prompt for every value listed above.
2. Load mappings.jsonl into memory, indexed by normalised clause key.
3. Optionally load the chunks JSONL, indexed by (doc_id, chunk_index).
4. Read gap.txt and suggest.txt from the templates directory.
5. CREATE TABLE IF NOT EXISTS suggestions (exact shared schema, see
   below) in the standards database.
6. Read clause_changes rows (oldest first by detected_at, up to
   batch_limit). A row is "pending" if no suggestions row exists yet
   for its change_id; see IDEMPOTENCY.
7. For each pending change, find affected documents via the mappings
   index (see MATCHING below), grouped by document: one suggestion row
   per (change_id, document_id, clause_id), with all matching sections
   of that document combined into a single gap prompt.
8. For each affected document: fill gap.txt with the new clause (old
   and new text, standard, clause_no, version, change type) and the
   document section(s), POST it to the local model, parse the impact
   rating (high/medium/low) from the gap output, then fill suggest.txt
   with that gap analysis and POST it again for the redline wording.
9. INSERT one suggestions row per affected document with parameterised
   SQL (psycopg %s placeholders), status defaulting to 'pending'.
   Commit after each change so a later failure does not lose earlier
   work. Print a counts summary at the end.

MATCHING (mappings.jsonl clause_key -> clause_changes row)
----------------------------------------------------------
mappings.jsonl clause_keys look like 'FDA eCFR:21 CFR 211.100',
'ICH:Q7 5.2' or 'EU GMP:Annex 11', while clause_changes rows carry
separate standard / chapter / clause_no fields, so there is no single
shared key. A mapping row matches a change when, after lower-casing
and whitespace normalisation, either:
  (a) the clause_no appears in the clause_key text after the ':', or
  (b) the full 'standard:clause_no' key equals the clause_key, or
  (c) the standard name appears in the clause_key prefix and the
      clause_no appears in the clause_key text.
The suggestions clause_id is written as '<standard> <clause_no>' and
document_id is the mapping row's doc_id.

IDEMPOTENCY
-----------
Before generating anything for a change, the script runs
SELECT ... FROM suggestions WHERE change_id = %s. If ANY row exists
for that change_id, the whole change is skipped -- no new model calls,
no new rows. Re-running the script therefore only processes changes
that have never produced a suggestion. (Consequence, by design: a
change that was partially processed in an interrupted run is NOT
backfilled on re-run; its existing rows stand and the change is
reported as already-processed.)

FAILURE HANDLING (no invented output)
-------------------------------------
- If the gap output contains no recognisable impact rating
  (high/medium/low), or the suggestion output is empty, or either
  model call fails, that document is skipped, counted as failed, and
  NO row is written for it. A rating is never guessed or defaulted.
- Malformed JSONL lines are skipped with a warning count, not fatal.
- A database or template error aborts the run before any model call
  for that stage; rows already committed for earlier changes stand.

CONFIG / DEFAULTS (every default below is an example -- replace with
your value)
-----------------------------------------------
- db_host default: localhost (example -- replace with your value)
- db_port default: 5432 (example -- replace with your value)
- db_name default: standards_db (example -- replace with your value)
- db_user default: app_user (example -- replace with your value)
- db_password default: none -- PGPASSWORD env or getpass only; there
  is no default password (example -- replace with your value applies
  to the env/getpass value you supply, never to a baked-in secret)
- mappings_path default: /data/staging/mappings.jsonl
  (example -- replace with your value)
- chunks_path default: (blank = no section text supplied)
  (example -- replace with your value)
- templates_directory default: /opt/airag/prompts
  (example -- replace with your value)
- model_endpoint default: http://localhost:11434/api/generate
  (example -- replace with your value)
- model_name default: local-model (example -- replace with your value)
- model_api_style default: ollama (example -- replace with your value)
- model_timeout_seconds default: 120
  (example -- replace with your value)
- detected_since default: (blank = no date window)
  (example -- replace with your value)
- batch_limit default: 20 (example -- replace with your value)
- include_needs_review_mappings default: no
  (example -- replace with your value)

SUGGESTIONS SCHEMA (exact, shared with the app worker; created with
CREATE TABLE IF NOT EXISTS before any read of clause_changes)
--------------------------------------------------------------
suggestions(suggestion_id SERIAL PRIMARY KEY,
            change_id INTEGER,
            document_id TEXT,
            clause_id TEXT,
            impact_rating TEXT CHECK (impact_rating IN
                ('high', 'medium', 'low')),
            suggested_text TEXT,
            status TEXT DEFAULT 'pending',
            reviewer TEXT,
            comment TEXT,
            created_at TIMESTAMPTZ DEFAULT now())

DEPENDENCIES
------------
- Python 3 standard library only, plus psycopg (psycopg v3) for the
  database connection. All SQL values use psycopg %s placeholders;
  no value is ever interpolated into SQL text.
- A PostgreSQL standards database with the clause_changes table
  (04-create-clause-registry.sql) already created.
- A reachable LOCAL model endpoint. No network call is made to any
  external model provider; the endpoint URL is supplied at runtime.
- gap.txt / suggest.txt templates (05-create-prompt-templates.py),
  using the placeholders {new_clause}, {doc_section} and
  {gap_analysis} respectively. Substitution is done with str.replace
  (not str.format) so other braces in a template cannot break it.

PREREQUISITES
-------------
- Clause changes recorded by 04-detect-standard-changes.py.
- mappings.jsonl from 03-map-documents-to-clauses.py.
"""
import getpass
import json
import os
import re
import sys
import urllib.error
import urllib.request
from datetime import datetime

try:
    import psycopg
except ImportError:  # pragma: no cover - dependency message only
    print("ERROR: psycopg is not installed. Install it (psycopg[binary]) "
          "and re-run. No database work was attempted.")
    sys.exit(1)

CREATE_SUGGESTIONS_SQL = """
CREATE TABLE IF NOT EXISTS suggestions (
    suggestion_id SERIAL PRIMARY KEY,
    change_id INTEGER,
    document_id TEXT,
    clause_id TEXT,
    impact_rating TEXT CHECK (impact_rating IN ('high', 'medium', 'low')),
    suggested_text TEXT,
    status TEXT DEFAULT 'pending',
    reviewer TEXT,
    comment TEXT,
    created_at TIMESTAMPTZ DEFAULT now()
)
"""

RATING_PATTERN = re.compile(
    r"impact\s+rating[^A-Za-z]*?(high|medium|low)", re.IGNORECASE)
RATING_FALLBACK = re.compile(r"\b(high|medium|low)\b", re.IGNORECASE)


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_optional(label, default=""):
    return input(f"{label} [{default}]: ").strip() or default


def prompt_secret_from_env_or_getpass(label):
    """Database password: PGPASSWORD env only, else getpass.

    Never argv, never echoed, never printed, never logged, never saved.
    """
    env_value = os.environ.get("PGPASSWORD", "").strip()
    if env_value:
        return env_value
    while True:
        value = getpass.getpass(f"{label}: ").strip()
        if value:
            return value
        print("A value is required.")


def prompt_int(label, default):
    while True:
        raw = input(f"{label} [{default}]: ").strip() or str(default)
        try:
            value = int(raw)
            if value > 0:
                return value
        except ValueError:
            pass
        print("Enter a whole number greater than 0.")


def prompt_yes_no(label, default):
    while True:
        value = (input(f"{label} [{default}]: ").strip() or default).lower()
        if value in ("yes", "no"):
            return value
        print("Enter yes or no.")


def prompt_date_or_blank(label):
    while True:
        raw = input(f"{label} []: ").strip()
        if not raw:
            return None
        try:
            datetime.strptime(raw, "%Y-%m-%d")
            return raw
        except ValueError:
            print("Enter a date as YYYY-MM-DD, or leave blank for no window.")


def normalise(text):
    return re.sub(r"\s+", " ", (text or "").strip().lower())


def mapping_matches_change(clause_key, standard, clause_no):
    """True if a mappings.jsonl clause_key refers to this clause change."""
    key = normalise(clause_key)
    if not key or not clause_no:
        return False
    clause_no_norm = normalise(clause_no)
    # Text after the 'standard:' prefix, e.g. '21 cfr 211.100'.
    key_body = key.split(":", 1)[1] if ":" in key else key
    key_prefix = key.split(":", 1)[0] if ":" in key else ""
    if key == normalise(f"{standard}:{clause_no}"):
        return True
    if clause_no_norm and clause_no_norm in key_body:
        return True
    if normalise(standard) and normalise(standard) in key_prefix \
            and clause_no_norm in key_body:
        return True
    return False


def load_mappings(path, include_needs_review):
    """Return {doc_id: [(clause_key, chunk_index)]} plus skip count."""
    by_doc = {}
    skipped = 0
    with open(path, encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                record = json.loads(line)
                doc_id = record["doc_id"]
                clause_key = record["clause_key"]
                status = record.get("status", "")
            except (ValueError, KeyError):
                skipped += 1
                continue
            if status == "needs_review" and not include_needs_review:
                continue
            by_doc.setdefault(doc_id, []).append(
                (clause_key, record.get("chunk_index")))
    return by_doc, skipped


def load_chunks(path):
    """Return {(doc_id, chunk_index): text} plus skip count."""
    chunks = {}
    skipped = 0
    with open(path, encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                record = json.loads(line)
                chunks[(record["doc_id"], record.get("chunk_index"))] = \
                    record.get("text", "")
            except (ValueError, KeyError):
                skipped += 1
    return chunks, skipped


def load_template(templates_directory, filename):
    path = os.path.join(templates_directory, filename)
    with open(path, encoding="utf-8") as handle:
        return handle.read()


def call_local_model(endpoint, api_style, model_name, prompt_text, timeout):
    """POST one prompt to the local model; return its text reply."""
    if api_style == "openai":
        payload = {"model": model_name,
                   "messages": [{"role": "user", "content": prompt_text}],
                   "stream": False}
        reply_field = ("choices", "message", "content")
    else:  # ollama
        payload = {"model": model_name, "prompt": prompt_text,
                   "stream": False}
        reply_field = ("response",)
    request = urllib.request.Request(
        endpoint, data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = json.loads(response.read().decode("utf-8"))
    except urllib.error.URLError as exc:
        raise RuntimeError(f"local model request failed: {exc}") from exc
    except ValueError as exc:
        raise RuntimeError("local model returned a non-JSON reply") from exc
    value = body
    for field in reply_field:
        if isinstance(value, list):
            if not value:
                raise RuntimeError("local model returned an empty reply")
            value = value[0]
        elif isinstance(value, dict):
            value = value.get(field)
        else:
            raise RuntimeError("local model reply had an unexpected shape")
        if value is None:
            raise RuntimeError("local model reply had an unexpected shape")
    return str(value).strip()


def parse_impact_rating(gap_output):
    """Extract high/medium/low from gap output; None if not stated."""
    match = RATING_PATTERN.search(gap_output)
    if match:
        return match.group(1).lower()
    # Fall back to the first bare rating word in the opening lines only,
    # so a passing mention deep in the text cannot invent the rating.
    opening = "\n".join(gap_output.splitlines()[:5])
    match = RATING_FALLBACK.search(opening)
    if match:
        return match.group(1).lower()
    return None


def main():
    print("=== Step 1: Suggestion settings ===")
    print("NOTE: this script only writes suggestions for human review.")
    print("It NEVER edits a controlled document.")
    db_host = prompt("Database host", "localhost")
    db_port = prompt("Database port", "5432")
    db_name = prompt("Database name", "standards_db")
    db_user = prompt("Database user", "app_user")
    db_password = prompt_secret_from_env_or_getpass(
        "Database password (PGPASSWORD env or hidden prompt; never echoed)")
    mappings_path = prompt("Mappings JSONL", "/data/staging/mappings.jsonl")
    chunks_path = prompt_optional(
        "Internal chunks JSONL (blank = no section text supplied)", "")
    templates_directory = prompt("Templates directory", "/opt/airag/prompts")
    model_endpoint = prompt("Local model endpoint",
                            "http://localhost:11434/api/generate")
    model_name = prompt("Local model name", "local-model")
    model_api_style = prompt("Local model API style (ollama/openai)",
                             "ollama").lower()
    if model_api_style not in ("ollama", "openai"):
        print("ERROR: model API style must be 'ollama' or 'openai'.")
        sys.exit(1)
    model_timeout = prompt_int("Local model timeout (seconds)", 120)
    detected_since = prompt_date_or_blank(
        "Only changes detected on/after YYYY-MM-DD (blank = all pending)")
    batch_limit = prompt_int("Maximum changes to process this run", 20)
    include_needs_review = prompt_yes_no(
        "Include 'needs_review' (keyword-only) mappings? (yes/no)", "no") == "yes"

    print("=== Step 2: Loading mappings and templates ===")
    try:
        mappings_by_doc, mappings_skipped = load_mappings(
            mappings_path, include_needs_review)
    except OSError as exc:
        print(f"ERROR: cannot read mappings file: {exc}")
        sys.exit(1)
    chunks = {}
    chunks_skipped = 0
    if chunks_path:
        try:
            chunks, chunks_skipped = load_chunks(chunks_path)
        except OSError as exc:
            print(f"ERROR: cannot read chunks file: {exc}")
            sys.exit(1)
    try:
        gap_template = load_template(templates_directory, "gap.txt")
        suggest_template = load_template(templates_directory, "suggest.txt")
    except OSError as exc:
        print(f"ERROR: cannot read prompt template: {exc}")
        sys.exit(1)
    print(f"  {len(mappings_by_doc)} documents in mappings "
          f"({mappings_skipped} malformed lines skipped).")
    if chunks_path:
        print(f"  {len(chunks)} chunks loaded "
              f"({chunks_skipped} malformed lines skipped).")

    print("=== Step 3: Connecting and preparing the suggestions table ===")
    try:
        conn = psycopg.connect(host=db_host, port=db_port, dbname=db_name,
                               user=db_user, password=db_password)
    except Exception as exc:  # connection errors only; password never printed
        print(f"ERROR: database connection failed: {exc}")
        sys.exit(1)
    # Deliberately drop the in-memory password reference once connected.
    db_password = None
    try:
        with conn.cursor() as cur:
            cur.execute(CREATE_SUGGESTIONS_SQL)
            if detected_since is not None:
                cur.execute(
                    "SELECT change_id, standard, chapter, clause_no, "
                    "old_version, new_version, change_type, old_text, "
                    "new_text, detected_at FROM clause_changes "
                    "WHERE detected_at >= %s "
                    "ORDER BY detected_at, change_id LIMIT %s",
                    (detected_since, batch_limit))
            else:
                cur.execute(
                    "SELECT change_id, standard, chapter, clause_no, "
                    "old_version, new_version, change_type, old_text, "
                    "new_text, detected_at FROM clause_changes "
                    "ORDER BY detected_at, change_id LIMIT %s",
                    (batch_limit,))
            changes = cur.fetchall()
        conn.commit()
    except Exception as exc:
        conn.rollback()
        print(f"ERROR: database read failed: {exc}")
        conn.close()
        sys.exit(1)
    print(f"  {len(changes)} clause change(s) in scope.")

    print("=== Step 4: Generating suggestions (review only) ===")
    already_processed = 0
    processed = 0
    written = 0
    no_mapping = 0
    failed = 0
    for (change_id, standard, chapter, clause_no, old_version, new_version,
         change_type, old_text, new_text, _detected_at) in changes:
        # Idempotency: any existing suggestion for this change skips it.
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM suggestions WHERE change_id = %s "
                        "LIMIT 1", (change_id,))
            if cur.fetchone() is not None:
                already_processed += 1
                print(f"  change {change_id}: already processed, skipping.")
                continue
        clause_id = f"{standard} {clause_no}"
        # Affected documents: every mapped doc whose clause_key matches
        # this change, grouped so each document gets exactly one row.
        affected = {}  # doc_id -> list of section texts
        for doc_id, entries in mappings_by_doc.items():
            sections = []
            for clause_key, chunk_index in entries:
                if not mapping_matches_change(clause_key, standard, clause_no):
                    continue
                if chunks_path:
                    section = chunks.get((doc_id, chunk_index), "")
                    sections.append(
                        section if section else
                        f"[No stored text for chunk {chunk_index}]")
                else:
                    sections.append(
                        f"[Section text not supplied: document {doc_id}, "
                        f"chunk {chunk_index}, mapped clause {clause_key}]")
            if sections:
                affected[doc_id] = sections
        if not affected:
            no_mapping += 1
            print(f"  change {change_id} ({clause_id}): "
                  "no mapped documents, nothing to suggest.")
            continue
        new_clause_text = (
            f"Standard: {standard}\nChapter: {chapter}\n"
            f"Clause: {clause_no}\nChange type: {change_type}\n"
            f"Old version: {old_version}\nNew version: {new_version}\n"
            f"Old text:\n{old_text or '[none: clause added]'}\n"
            f"New text:\n{new_text or '[none: clause removed]'}")
        rows_for_change = []
        for doc_id, sections in sorted(affected.items()):
            doc_section = "\n---\n".join(sections)
            gap_prompt = gap_template.replace("{new_clause}", new_clause_text)
            gap_prompt = gap_prompt.replace("{doc_section}", doc_section)
            try:
                gap_output = call_local_model(
                    model_endpoint, model_api_style, model_name,
                    gap_prompt, model_timeout)
                rating = parse_impact_rating(gap_output)
                if rating is None:
                    raise RuntimeError(
                        "gap output stated no impact rating")
                suggest_prompt = suggest_template.replace(
                    "{gap_analysis}", gap_output)
                suggested_text = call_local_model(
                    model_endpoint, model_api_style, model_name,
                    suggest_prompt, model_timeout)
                if not suggested_text:
                    raise RuntimeError("suggestion output was empty")
            except RuntimeError as exc:
                failed += 1
                print(f"  change {change_id}, doc {doc_id}: "
                      f"skipped ({exc}); no row written.")
                continue
            rows_for_change.append((change_id, doc_id, clause_id, rating,
                                    suggested_text))
            print(f"  change {change_id}, doc {doc_id}: "
                  f"impact {rating}, suggestion ready.")
        if rows_for_change:
            try:
                with conn.cursor() as cur:
                    cur.executemany(
                        "INSERT INTO suggestions (change_id, document_id, "
                        "clause_id, impact_rating, suggested_text) "
                        "VALUES (%s, %s, %s, %s, %s)", rows_for_change)
                conn.commit()
            except Exception as exc:
                conn.rollback()
                failed += len(rows_for_change)
                print(f"  change {change_id}: database write failed: {exc}; "
                      "no rows written for this change.")
                continue
            processed += 1
            written += len(rows_for_change)
    conn.close()

    print("=== Step 5: Summary ===")
    print(f"Changes in scope: {len(changes)}")
    print(f"Already processed (skipped): {already_processed}")
    print(f"Changes with suggestions written: {processed}")
    print(f"Suggestion rows written (status 'pending'): {written}")
    print(f"Changes with no mapped documents: {no_mapping}")
    print(f"Documents skipped (model/rating failure, no row): {failed}")
    print("Done. Suggestions await human review; no controlled document "
          "was modified.")


if __name__ == "__main__":
    main()

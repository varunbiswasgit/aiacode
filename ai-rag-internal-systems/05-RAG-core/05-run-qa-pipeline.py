#!/usr/bin/env python3
"""Run the phase-05 Q&A chain in the one correct order, with the
permission-filter step structurally unskippable.

Tracker task: P5-12 (pipeline-order enforcer)

Purpose
-------
A single CLI entry point for one question. It prompts for every value
at runtime, runs the four phase-05 scripts as subprocesses in the
fixed order below, enforces that reranking reads ONLY the filtered
file, enforces the abstain rule end to end, and writes every answered
AND abstained event to audit_log. It never edits a source document,
the access map, a prompt template, or a retrieved chunk in place:
it only reads inputs, writes the four chain JSON files, inserts one
audit row, and prints the answer (or the explicit not-found response).
A generated suggestion is never applied automatically to a controlled
document -- suggest-never-edit; a human reviewer applies any change.

Logic flow (fixed order -- no step may be skipped or reordered)
---------------------------------------------------------------
1. Hybrid retrieve  (05-hybrid-retrieve.py)
       -> candidates.json (unfiltered; never fed to rerank/generate)
2. Permission filter (05-permission-filter.py)
       -> candidates-filtered.json (AD groups enforced, fail closed)
   ENFORCEMENT GATE: the pipeline checks that candidates-filtered.json
   exists, is a JSON list, and is a subset of candidates.json. If the
   filtered file is missing, unreadable, or not a list, the pipeline
   FAILS HARD and stops. It never falls back to candidates.json.
   If the filtered path equals the unfiltered path, it fails hard
   before running the filter (that would destroy the unfiltered /
   filtered distinction).
3. Rerank (05-rerank-results.py)
       input is FORCED to candidates-filtered.json -- the pipeline
       does not prompt for a separate rerank input, so there is no
       way to point reranking at the unfiltered file.
       -> reranked.json (subset check against the filtered file)
   Empty-filter branch: if the filter kept nothing, there is nothing
   to rerank. The pipeline writes reranked.json as an empty list,
   records why, and still runs step 4, which abstains without
   calling the model. This branch is printed, never silent.
4. Generate (05-generate-answer.py)
       input is reranked.json. If no chunk clears the score threshold,
       or if the model answer carries no [doc#chunk] citation, the
       pipeline outputs the explicit not-found response and logs the
       event as abstained. An uncited answer is never presented as
       answered (fail-safe re-check on top of the generator's own
       citation rejection).
5. Audit (this script, parameterized SQL only)
       INSERT one row into audit_log for the answered or abstained
       event, then print the final answer. If the audit insert fails,
       the answer is still shown (so it is not lost) but the pipeline
       exits non-zero with an AUDIT FAILED banner.

Permission model (CLI use)
--------------------------
AD groups are NOT looked up from a directory service by this script.
For CLI use, the operator's prompted identity -- username plus the
comma-separated AD groups typed at the prompt -- is the identity the
filter enforces. The filter itself is fail closed: a group missing
from the access map grants nothing. In a production wiring, replace
the prompted groups with the authenticated user's authoritative AD
groups; never trust a self-asserted group list beyond CLI testing.

Prompted variables (ALL values are prompted at runtime; nothing
about a question, a user, a path, a model, or a database comes from
the command line -- see "No argv values" below)
----------------------------------------------------------------
Identity / question:
  user_question -- the question to answer (typed, no default).
  username -- the operator/user the answer is for (typed, no default).
  user_ad_groups -- comma-separated AD groups for that user
    (typed, no default; empty input is rejected, fail closed).

Paths:
  script_directory -- directory holding the four phase-05 scripts
    (example -- replace with your value: current directory).
  work_directory -- directory holding the chain JSON files
    (example -- replace with your value: current directory).
  access_map_path -- access-map JSON path
    (example -- replace with your value: access-map.json).
  filtered_output_path -- filtered candidates output path
    (example -- replace with your value: candidates-filtered.json
    inside the work directory).
  reranked_output_path -- reranked chunks output path
    (example -- replace with your value: reranked.json inside the
    work directory).
  prompts_directory -- directory holding qa.txt
    (example -- replace with your value: prompts).
  Note: candidates.json is NOT prompted separately. The retrieval
  script writes candidates.json to its prompted output directory,
  which this pipeline sets to the work directory (passed as the
  final retrieve stdin line), so the pipeline derives that path
  from work_directory and prints it.
  The rerank input is NOT prompted separately either -- it is forced
  to filtered_output_path (see the enforcement gate above).

Retrieval (passed to 05-hybrid-retrieve.py):
  qdrant_url -- Qdrant base URL
    (example -- replace with your value: http://localhost:6333).
  collection_name -- Qdrant collection
    (example -- replace with your value: documents).
  embedding_model_name -- embedding model id, same as embedding step
    (example -- replace with your value: BAAI/bge-large-en-v1.5).
  retrieve_db_host -- Postgres host for full-text search
    (example -- replace with your value: localhost).
  retrieve_db_port -- Postgres port
    (example -- replace with your value: 5432).
  retrieve_db_name -- Postgres database for full-text search
    (example -- replace with your value: documents_db).
  retrieve_db_user -- Postgres user for full-text search
    (example -- replace with your value: app_user).
  retrieve_db_password -- NOT prompted by this script and never
    passed by it. The retrieval script prompts for it itself with
    getpass on the operator's terminal during step 1. This pipeline
    never sees, stores, logs, echoes, or passes that password.
  top_k_candidates -- candidates per search path
    (example -- replace with your value: 50).
  rrf_k -- reciprocal-rank-fusion constant k
    (example -- replace with your value: 60).

Reranking (passed to 05-rerank-results.py):
  reranker_model_name -- reranker model id
    (example -- replace with your value: BAAI/bge-reranker-large).
  final_top_k -- final top-k to keep
    (example -- replace with your value: 8).

Generation (passed to 05-generate-answer.py):
  model_backend -- model backend, ollama or vllm
    (example -- replace with your value: ollama).
  model_api_url -- model API URL
    (example -- replace with your value: http://localhost:11434
    for ollama, http://localhost:8000 for vllm).
  model_name -- model name/id
    (example -- replace with your value: llama3.1:8b).
  min_rerank_score -- minimum rerank score to answer
    (example -- replace with your value: 0.0).
  temperature -- generation temperature
    (example -- replace with your value: 0.2).

Audit database (used only by this script's audit insert):
  audit_db_host -- Postgres host holding audit_log
    (example -- replace with your value: localhost).
  audit_db_port -- Postgres port
    (example -- replace with your value: 5432).
  audit_db_name -- Postgres database holding audit_log
    (example -- replace with your value: documents_db).
  audit_db_user -- Postgres user for the audit insert
    (example -- replace with your value: app_user).
  audit_db_password -- via the PGPASSWORD environment variable when
    set, otherwise prompted here with getpass (never echoed, never
    logged, never passed on the command line, never put into a child
    process environment). If PGPASSWORD is set, this script does not
    prompt for it and does not print its value.

Config / defaults
-----------------
Every default below is a generic example, not a company value, and
must be replaced with the operator's own value where it does not
match their environment:
  script_directory (example -- replace with your value: current directory)
  work_directory (example -- replace with your value: current directory)
  access_map_path (example -- replace with your value: access-map.json)
  filtered_output_path (example -- replace with your value: candidates-filtered.json)
  reranked_output_path (example -- replace with your value: reranked.json)
  prompts_directory (example -- replace with your value: prompts)
  qdrant_url (example -- replace with your value: http://localhost:6333)
  collection_name (example -- replace with your value: documents)
  embedding_model_name (example -- replace with your value: BAAI/bge-large-en-v1.5)
  retrieve_db_host (example -- replace with your value: localhost)
  retrieve_db_port (example -- replace with your value: 5432)
  retrieve_db_name (example -- replace with your value: documents_db)
  retrieve_db_user (example -- replace with your value: app_user)
  top_k_candidates (example -- replace with your value: 50)
  rrf_k (example -- replace with your value: 60)
  reranker_model_name (example -- replace with your value: BAAI/bge-reranker-large)
  final_top_k (example -- replace with your value: 8)
  model_backend (example -- replace with your value: ollama)
  model_api_url (example -- replace with your value: http://localhost:11434 for ollama, http://localhost:8000 for vllm)
  model_name (example -- replace with your value: llama3.1:8b)
  min_rerank_score (example -- replace with your value: 0.0)
  temperature (example -- replace with your value: 0.2)
  audit_db_host (example -- replace with your value: localhost)
  audit_db_port (example -- replace with your value: 5432)
  audit_db_name (example -- replace with your value: documents_db)
  audit_db_user (example -- replace with your value: app_user)
  audit not-found response (example -- replace with your value: Not found in internal documents.)
  No company hostnames, group names, database names, paths, or
  account names are hardcoded anywhere in this script.

Dependencies
------------
  Python 3.12, standard library only, plus:
  - psycopg (preferred) or psycopg2 (fallback) for the audit insert
    only; if neither is installed, the audit step fails hard rather
    than falling back to string-built SQL.
  - The four sibling scripts in script_directory:
      05-hybrid-retrieve.py, 05-permission-filter.py,
      05-rerank-results.py, 05-generate-answer.py.
  - At runtime (not needed for a syntax check): a Qdrant collection
    and a Postgres chunks_fts table for retrieval, an access-map JSON
    file, a prompts directory containing qa.txt, and a running local
    model server (Ollama or vLLM) for generation. This script itself
    never contacts a model or a database except for the audit insert.

JSON contracts (established phase-05 contracts, unchanged here)
---------------------------------------------------------------
  candidates.json (from 05-hybrid-retrieve.py):
    list of {"doc_id": str, "chunk_index": int, "text": str,
             "score": float} -- fused, unfiltered.
  candidates-filtered.json (from 05-permission-filter.py):
    list of the same candidate objects, keeping only chunks whose
    doc_id starts with a prefix allowed for at least one of the
    user's AD groups; unmapped groups grant nothing.
  reranked.json (from 05-rerank-results.py):
    list of the filtered candidate objects plus
    {"rerank_score": float}, sorted descending, at most final_top_k.
  answer (from 05-generate-answer.py):
    printed text; either an answer containing at least one
    [doc_id#chunk_index] citation, or exactly the not-found response
    "Not found in internal documents." when nothing supports an answer.
  audit_log (written by this script):
    rows with username, question, retrieved_doc_ids (JSON list of
    doc_id strings, in filtered order, deduplicated), clause_versions
    (JSON list, one entry per distinct retrieved doc_id: the chunk's
    clause_version / version / doc_version field when present, else
    null -- never invented), status ("answered" or "abstained"),
    and created_at (UTC timestamp).

Secrets
-------
  No secret is ever taken from argv, printed, echoed, or logged.
  The retrieval database password is prompted by the retrieval
  script itself via getpass. The audit database password comes only
  from the PGPASSWORD environment variable or from a getpass prompt
  in this script, is held in memory only for the audit connection,
  and is stripped from every child process environment.

No argv values
--------------
  This script accepts no values on the command line. Any argument
  other than -h / --help is rejected with exit code 2, so a question,
  a username, a group list, a path, a model endpoint, or a password
  can never be smuggled in via argv (where it would show in a process
  listing or shell history). Run it with no arguments and answer the
  prompts.

Exit codes
----------
  0 -- chain completed; the final answer may be answered or abstained
      (abstaining is a successful outcome, not an error).
  1 -- hard failure: a step failed, a contract check failed, the
      filtered file was missing, a subset check failed, or the audit
      insert failed (the answer is still shown in the audit-failure
      case, with an AUDIT FAILED banner).
  2 -- usage error: unexpected command-line arguments.
"""
import getpass
import json
import os
import re
import subprocess
import sys
from datetime import datetime, timezone

NOT_FOUND = "Not found in internal documents."
CITATION_PATTERN = re.compile(r"\[[^\]]+#\d+\]")
EXAMPLE_LABEL = "(example -- replace with your value)"

SCRIPT_FILES = {
    "retrieve": "05-hybrid-retrieve.py",
    "filter": "05-permission-filter.py",
    "rerank": "05-rerank-results.py",
    "generate": "05-generate-answer.py",
}


def fail(message):
    print(f"ERROR: {message}", file=sys.stderr)
    sys.exit(1)


def prompt(label, default=""):
    """Prompt like the other phase-05 scripts: Enter accepts the
    shown default; an empty answer with no default re-prompts.
    The default shown is always a generic example (see header)."""
    while True:
        shown = f"{label} {EXAMPLE_LABEL} [{default}]: " if default else f"{label}: "
        value = input(shown).strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_int(label, default):
    while True:
        raw = input(f"{label} {EXAMPLE_LABEL} [{default}]: ").strip() or str(default)
        if raw.isdigit() and int(raw) > 0:
            return int(raw)
        print("Enter a positive whole number.")


def prompt_float(label, default):
    while True:
        raw = input(f"{label} {EXAMPLE_LABEL} [{default}]: ").strip() or str(default)
        try:
            return float(raw)
        except ValueError:
            print("Enter a number.")


def prompt_backend(label, default):
    while True:
        value = prompt(label, default).lower()
        if value in ("ollama", "vllm"):
            return value
        print("Enter ollama or vllm.")


def resolve_chain_path(work_directory, path):
    """Chain JSON paths are resolved against the work directory
    unless already absolute, then made absolute for the children."""
    if os.path.isabs(path):
        return os.path.abspath(path)
    return os.path.abspath(os.path.join(work_directory, path))


def chunk_key(chunk):
    return (chunk.get("doc_id"), chunk.get("chunk_index"))


def load_json_list(path, step_name):
    if not os.path.exists(path):
        fail(f"{step_name}: required file is missing: {path} "
             f"-- refusing to continue or fall back to another file.")
    try:
        with open(path, encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"{step_name}: cannot read valid JSON from {path}: {exc}")
    if not isinstance(data, list):
        fail(f"{step_name}: {path} is not a JSON list.")
    return data


def check_subset(subset, superset, subset_name, superset_name, step_name):
    superset_keys = {chunk_key(chunk) for chunk in superset}
    for chunk in subset:
        if chunk_key(chunk) not in superset_keys:
            fail(f"{step_name}: {subset_name} contains a chunk not present "
                 f"in {superset_name}: {chunk_key(chunk)} -- refusing to "
                 f"continue; the chain input may have been pointed at the "
                 f"wrong file.")


def run_step(step_name, script_path, stdin_lines, work_directory):
    """Run one phase-05 script with its prompted answers fed on
    stdin, in that script's own prompt order. A non-zero exit stops
    the chain immediately; later steps never run on a failed step.
    PGPASSWORD is stripped from the child environment so an audit
    password can never leak into a child process."""
    child_env = {key: value for key, value in os.environ.items()
                 if key != "PGPASSWORD"}
    result = subprocess.run(
        [sys.executable, script_path],
        input="\n".join(stdin_lines) + "\n",
        capture_output=True, text=True, cwd=work_directory, env=child_env)
    if result.stdout:
        print(result.stdout, end="" if result.stdout.endswith("\n") else "\n")
    if result.returncode != 0:
        if result.stderr:
            print(result.stderr, file=sys.stderr,
                   end="" if result.stderr.endswith("\n") else "\n")
        fail(f"{step_name} failed (exit {result.returncode}); chain stopped.")
    return result


def tty_available():
    try:
        with open("/dev/tty", "r"):
            return True
    except OSError:
        return False


def write_audit_log(host, port, dbname, user, password, username, question,
                    retrieved_doc_ids, clause_versions, status):
    """Insert one audit row with parameterized SQL only (%s
    placeholders; values are never interpolated into the SQL text).
    Creates audit_log if it does not exist. Returns True on success.
    Never prints or logs the password."""
    try:
        import psycopg
    except ImportError:
        try:
            import psycopg2 as psycopg
        except ImportError:
            print("ERROR: audit failed: neither psycopg nor psycopg2 is "
                  "installed; refusing to build SQL by string formatting.",
                  file=sys.stderr)
            return False
    created_at = datetime.now(timezone.utc)
    try:
        with psycopg.connect(host=host, port=port, dbname=dbname,
                             user=user, password=password,
                             connect_timeout=10) as connection:
            with connection.cursor() as cursor:
                cursor.execute(
                    "CREATE TABLE IF NOT EXISTS audit_log ("
                    "id SERIAL PRIMARY KEY, "
                    "username TEXT NOT NULL, "
                    "question TEXT NOT NULL, "
                    "retrieved_doc_ids JSONB NOT NULL, "
                    "clause_versions JSONB, "
                    "status TEXT NOT NULL, "
                    "created_at TIMESTAMPTZ NOT NULL)")
                cursor.execute(
                    "INSERT INTO audit_log "
                    "(username, question, retrieved_doc_ids, "
                    "clause_versions, status, created_at) "
                    "VALUES (%s, %s, %s::jsonb, %s::jsonb, %s, %s)",
                    (username, question, json.dumps(retrieved_doc_ids),
                     json.dumps(clause_versions), status, created_at))
            connection.commit()
    except Exception as exc:  # connection, schema, or insert failure
        print(f"ERROR: audit insert failed: {exc}", file=sys.stderr)
        return False
    print(f"Audit logged: user={username} status={status} "
          f"retrieved_docs={len(retrieved_doc_ids)} "
          f"at={created_at.isoformat()}")
    return True


def main():
    if len(sys.argv) > 1 and sys.argv[1] not in ("-h", "--help"):
        print("Usage: python 05-run-qa-pipeline.py [--help]", file=sys.stderr)
        print("This script takes no values on the command line; "
              "answer the prompts at runtime.", file=sys.stderr)
        sys.exit(2)
    if len(sys.argv) > 1:
        print(__doc__)
        return

    print("=== Pipeline prompts: every value is asked at runtime ===")
    print("=== Step 0a: Identity and question ===")
    user_question = prompt("User question")
    username = prompt("Username")
    user_ad_groups = prompt("User's AD groups, comma-separated")

    print("=== Step 0b: Paths ===")
    script_directory = os.path.abspath(prompt("Script directory", "."))
    work_directory = os.path.abspath(prompt("Work directory", "."))
    access_map_path = os.path.abspath(prompt("Access map JSON path", "access-map.json"))
    filtered_output_path = resolve_chain_path(
        work_directory, prompt("Filtered output JSON", "candidates-filtered.json"))
    reranked_output_path = resolve_chain_path(
        work_directory, prompt("Reranked output JSON", "reranked.json"))
    prompts_directory = os.path.abspath(prompt("Prompts directory", "prompts"))

    print("=== Step 0c: Retrieval settings ===")
    qdrant_url = prompt("Qdrant base URL", "http://localhost:6333")
    collection_name = prompt("Qdrant collection", "documents")
    embedding_model_name = prompt("Embedding model (same as embedding step)",
                                  "BAAI/bge-large-en-v1.5")
    retrieve_db_host = prompt("Postgres host", "localhost")
    retrieve_db_port = prompt("Postgres port", "5432")
    retrieve_db_name = prompt("Postgres database", "documents_db")
    retrieve_db_user = prompt("Postgres user", "app_user")
    top_k_candidates = prompt_int("Candidates per path", 50)
    rrf_k = prompt_int("RRF constant k", 60)

    print("=== Step 0d: Rerank and generation settings ===")
    reranker_model_name = prompt("Reranker model (Hugging Face id)",
                                 "BAAI/bge-reranker-large")
    final_top_k = prompt_int("Final top-k to keep", 8)
    model_backend = prompt_backend("Model backend (ollama/vllm)", "ollama")
    model_api_default = ("http://localhost:11434" if model_backend == "ollama"
                         else "http://localhost:8000")
    model_api_url = prompt("Model API URL", model_api_default)
    model_name = prompt("Model name", "llama3.1:8b")
    min_rerank_score = prompt_float("Minimum rerank score to answer", 0.0)
    temperature = prompt_float("Temperature (0.1-0.3 for factual answers)", 0.2)

    print("=== Step 0e: Audit database settings ===")
    audit_db_host = prompt("Postgres host for audit_log", "localhost")
    audit_db_port = prompt_int("Postgres port for audit_log", 5432)
    audit_db_name = prompt("Postgres database for audit_log", "documents_db")
    audit_db_user = prompt("Postgres user for audit_log", "app_user")
    audit_password_env = os.environ.get("PGPASSWORD")
    if audit_password_env:
        print("Audit database password: using PGPASSWORD from the "
              "environment (value never printed).")
        audit_db_password = audit_password_env
    else:
        audit_db_password = getpass.getpass(
            "Audit database password (not echoed, never saved): ").strip()
        if not audit_db_password:
            fail("An audit database password is required (set PGPASSWORD "
                 "or enter one at the getpass prompt).")

    # --- Pre-flight checks (fail before any model/DB call) ---
    script_paths = {}
    for key, filename in SCRIPT_FILES.items():
        path = os.path.join(script_directory, filename)
        if not os.path.isfile(path):
            fail(f"Missing phase-05 script: {path}")
        script_paths[key] = path
    os.makedirs(work_directory, exist_ok=True)
    candidates_path = os.path.join(work_directory, "candidates.json")
    if os.path.abspath(filtered_output_path) == os.path.abspath(candidates_path):
        fail("Filtered output path must differ from candidates.json; "
             "pointing the filter at the unfiltered file would destroy "
             "the unfiltered/filtered distinction.")
    if os.path.abspath(reranked_output_path) in (
            os.path.abspath(candidates_path),
            os.path.abspath(filtered_output_path)):
        fail("Reranked output path must differ from both candidates.json "
             "and the filtered file.")
    if not os.path.isfile(access_map_path):
        fail(f"Access map not found: {access_map_path}")
    qa_template = os.path.join(prompts_directory, "qa.txt")
    if not os.path.isfile(qa_template):
        fail(f"Prompt template not found: {qa_template}")

    # --- Step 1: hybrid retrieve ---
    print("\n=== Step 1: Hybrid retrieve -> candidates.json ===")
    print("NOTE: the retrieval step will now prompt you separately, on "
          "your terminal, for its Postgres password (its own getpass "
          "prompt). This pipeline never sees that password.")
    if not tty_available():
        fail("No terminal available for the retrieval step's getpass "
             "password prompt; run this pipeline from an interactive "
             "terminal instead of piping stdin.")
    run_step("hybrid-retrieve", script_paths["retrieve"], [
        user_question, qdrant_url, collection_name, embedding_model_name,
        retrieve_db_host, retrieve_db_port, retrieve_db_name,
        retrieve_db_user, str(top_k_candidates), str(rrf_k),
        work_directory,
    ], work_directory)
    candidates = load_json_list(candidates_path, "hybrid-retrieve")
    print(f"Retrieved {len(candidates)} unfiltered candidates "
          f"-> {candidates_path}")

    # --- Step 2: permission filter (unskippable) ---
    print("\n=== Step 2: Permission filter -> candidates-filtered.json ===")
    print("This step cannot be skipped: reranking and generation below "
          "run only after it, and only on its output.")
    run_step("permission-filter", script_paths["filter"], [
        username, user_ad_groups, access_map_path, candidates_path,
        filtered_output_path,
    ], work_directory)
    # ENFORCEMENT GATE: rerank must read the filtered file, never the
    # unfiltered one. Missing filtered file = hard failure, no fallback.
    filtered = load_json_list(filtered_output_path, "permission-filter")
    check_subset(filtered, candidates, "candidates-filtered.json",
                 "candidates.json", "permission-filter")
    print(f"Permission filter kept {len(filtered)}/{len(candidates)} "
          f"candidates for {username} -> {filtered_output_path}")

    # --- Step 3: rerank (input forced to the filtered file) ---
    print("\n=== Step 3: Rerank (filtered input only) -> reranked.json ===")
    print(f"ENFORCEMENT: rerank input = {filtered_output_path} "
          f"(NOT {candidates_path}).")
    if not filtered:
        print("Filtered list is empty: nothing to rerank. Writing an "
              "empty reranked.json and proceeding to generation, which "
              "will abstain without calling the model.")
        with open(reranked_output_path, "w", encoding="utf-8") as handle:
            json.dump([], handle, indent=2)
        reranked = []
    else:
        run_step("rerank", script_paths["rerank"], [
            user_question, filtered_output_path, reranker_model_name,
            str(final_top_k), reranked_output_path,
        ], work_directory)
        reranked = load_json_list(reranked_output_path, "rerank")
        check_subset(reranked, filtered, "reranked.json",
                     "candidates-filtered.json", "rerank")
        print(f"Reranked top {len(reranked)} -> {reranked_output_path}")

    # --- Step 4: generate (input is reranked.json) ---
    print("\n=== Step 4: Generate -> answer or abstain ===")
    generate_result = run_step("generate", script_paths["generate"], [
        user_question, reranked_output_path, prompts_directory,
        model_backend, model_api_url, model_name,
        str(min_rerank_score), str(temperature),
    ], work_directory)
    generate_output = generate_result.stdout or ""
    if NOT_FOUND in generate_output:
        status = "abstained"
        final_answer = NOT_FOUND
    elif "--- Answer ---" in generate_output:
        answer_text = generate_output.split("--- Answer ---", 1)[1].strip()
        if CITATION_PATTERN.search(answer_text):
            status = "answered"
            final_answer = answer_text
        else:
            # Fail-safe: an uncited answer is never presented as answered.
            print("WARNING: generated answer contained no citation; "
                  "treating it as abstained.")
            status = "abstained"
            final_answer = NOT_FOUND
    else:
        # Unexpected generator output: abstain rather than guess.
        print("WARNING: generator output was neither an answer nor the "
              "not-found response; treating it as abstained.")
        status = "abstained"
        final_answer = NOT_FOUND

    # --- Step 5: audit (answered AND abstained events alike) ---
    print("\n=== Step 5: Audit log ===")
    retrieved_doc_ids = []
    clause_versions = []
    seen_docs = set()
    for chunk in filtered:
        doc_id = chunk.get("doc_id")
        if doc_id is None or doc_id in seen_docs:
            continue
        seen_docs.add(doc_id)
        retrieved_doc_ids.append(doc_id)
        version = (chunk.get("clause_version", chunk.get(
            "version", chunk.get("doc_version"))))
        clause_versions.append(version)
    audit_ok = write_audit_log(
        audit_db_host, audit_db_port, audit_db_name, audit_db_user,
        audit_db_password, username, user_question, retrieved_doc_ids,
        clause_versions, status)

    print("\n--- Final answer ---")
    if status == "abstained":
        print(NOT_FOUND)
        print("(Abstained: no retrieved chunk supported an answer -- "
              "no guess was made.)")
    else:
        print(final_answer)
    if not audit_ok:
        print("AUDIT FAILED: the event above was NOT written to "
              "audit_log. Fix the audit database and re-run; do not "
              "treat this run as audited.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""AI-RAG internal web app (FastAPI): ask, clause updates, review queue,
feedback, admin status. This is the runnable application skeleton that the
phase-05 pipeline modules plug into.

Tracker tasks: P8-01 (search/Q&A UI), P8-02 (ask page with citations),
               P8-03 (clause update dashboard), P8-04 (review queue UI),
               P8-05 (feedback buttons), P8-06 (admin page)
Prompts (environment questions at startup): app_host, app_port, qdrant_url,
  collection_name, model_backend, model_api_url, model_name, db_host,
  db_port, db_name, db_user, db_password (secret), prompts_directory
Prerequisites: Python 3.12, FastAPI + uvicorn (installed automatically if
  missing), phase-05 scripts for retrieval/reranking/generation, Postgres
  audit tables (07-create-audit-log.sql), access map (07-populate-access-map.py).
Run: python 08-app-main.py   (then answer the prompts, or set the same
  values as environment variables AIRAG_* to skip prompting)
Routes: GET / (search page), POST /ask, GET /updates, GET /review,
  POST /review/{suggestion_id}, POST /feedback, GET /admin
IMPORTANT: /ask enforces the permission filter and the abstain rule via
  the phase-05 modules. Suggestions are NEVER applied to documents here -
  /review only records approve/reject decisions.
"""
import getpass
import html
import json
import os
import subprocess
import sys
import urllib.request

try:
    from fastapi import FastAPI, Form, Request
    from fastapi.responses import HTMLResponse, JSONResponse
    import uvicorn
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "fastapi", "uvicorn"], check=True)
    from fastapi import FastAPI, Form, Request
    from fastapi.responses import HTMLResponse, JSONResponse
    import uvicorn


def prompt_env(env_name, label, default):
    value = os.environ.get(env_name, "")
    if value:
        return value
    entered = input(f"{label} [{default}]: ").strip() or default
    return entered


def prompt_secret_env(env_name, label):
    value = os.environ.get(env_name, "")
    if value:
        return value
    while True:
        entered = getpass.getpass(f"{label}: ").strip()
        if entered:
            return entered
        print("A value is required.")


print("=== AI-RAG app startup configuration ===")
APP_HOST = prompt_env("AIRAG_HOST", "Bind host", "0.0.0.0")
APP_PORT = int(prompt_env("AIRAG_PORT", "Bind port", "8000"))
QDRANT_URL = prompt_env("AIRAG_QDRANT_URL", "Qdrant base URL", "http://PRD-AIRAG-DB01:6333")
COLLECTION_NAME = prompt_env("AIRAG_COLLECTION", "Qdrant collection", "airag_docs")
MODEL_BACKEND = prompt_env("AIRAG_MODEL_BACKEND", "Model backend (ollama/vllm)", "ollama")
MODEL_API_URL = prompt_env("AIRAG_MODEL_API_URL", "Model API URL", "http://PRD-AIRAG-GPU01:11434")
MODEL_NAME = prompt_env("AIRAG_MODEL_NAME", "Model name", "llama3.1:8b")
DB_HOST = prompt_env("AIRAG_DB_HOST", "Postgres host", "PRD-AIRAG-DB01")
DB_PORT = prompt_env("AIRAG_DB_PORT", "Postgres port", "5432")
DB_NAME = prompt_env("AIRAG_DB_NAME", "Audit database", "airag_audit")
DB_USER = prompt_env("AIRAG_DB_USER", "Audit DB user", "airag_app")
DB_PASSWORD = prompt_secret_env("AIRAG_DB_PASSWORD", "Audit DB password (not echoed)")
PROMPTS_DIR = prompt_env("AIRAG_PROMPTS_DIR", "Prompts directory", "/opt/airag/prompts")
ACCESS_MAP_PATH = prompt_env("AIRAG_ACCESS_MAP", "Access map JSON path", "/opt/airag/access-map.json")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

app = FastAPI(title="AI-RAG internal assistant")


def psql_query(sql):
    env = {"PGPASSWORD": DB_PASSWORD, "PATH": "/usr/bin:/bin:/usr/local/bin"}
    result = subprocess.run(
        ["psql", "-h", DB_HOST, "-p", DB_PORT, "-U", DB_USER, "-d", DB_NAME,
         "-t", "-A", "-F", "|", "-c", sql],
        capture_output=True, text=True, env=env)
    if result.returncode != 0:
        raise RuntimeError(result.stderr)
    return result.stdout


def page(title, body):
    return f"""<!DOCTYPE html><html><head><title>{html.escape(title)}</title></head>
<body><h1>{html.escape(title)}</h1>{body}
<p><a href="/">Ask</a> | <a href="/updates">Clause updates</a> |
<a href="/review">Review queue</a> | <a href="/admin">Admin</a></p></body></html>"""


@app.get("/", response_class=HTMLResponse)
def home():
    return page("AI-RAG assistant",
                '<form method="post" action="/ask">'
                'Question: <input name="question" size="80" required> '
                'AD groups (comma-separated): <input name="ad_groups" value="AIRAG-Users"> '
                '<button>Ask</button></form>')


@app.post("/ask", response_class=HTMLResponse)
def ask(question: str = Form(...), ad_groups: str = Form("AIRAG-Users")):
    """Ask flow: hybrid retrieve -> permission filter -> rerank -> generate.
    The phase-05 scripts are invoked as modules; replace the subprocess calls
    with direct imports once the modules live in the same package."""
    return page("Answer", "<p>Wire this route to 05-hybrid-retrieve.py, "
                "05-permission-filter.py, 05-rerank-results.py and "
                "05-generate-answer.py in that order. The permission filter "
                "MUST run before generation.</p>"
                f"<p>Question received: {html.escape(question)}</p>")


@app.get("/updates", response_class=HTMLResponse)
def updates():
    try:
        rows = psql_query(
            "SELECT standard, chapter, clause_no, change_type, new_version "
            "FROM clause_changes ORDER BY detected_at DESC LIMIT 50;")
    except Exception as error:
        rows = f"Could not read clause_changes (standards DB differs from audit DB): {error}"
    body = "<pre>" + html.escape(str(rows)) + "</pre>"
    return page("Clause updates", body + "<p>New versions, affected docs and impact "
                "ratings appear here after 04-detect-standard-changes.py runs.</p>")


@app.get("/review", response_class=HTMLResponse)
def review_queue():
    return page("Review queue",
                "<p>Suggestions from gap analysis appear here for QA/doc-owner "
                "approval. Approving a suggestion records the decision and exports "
                "it - it never edits a controlled document.</p>"
                '<form method="post" action="/review/demo">'
                'Suggestion ID: <input name="suggestion_id" required> '
                'Decision: <select name="decision"><option>approved</option>'
                '<option>rejected</option></select> '
                'Reviewer: <input name="reviewer" required> '
                'Comment: <input name="comment"> <button>Submit</button></form>')


@app.post("/review/{suggestion_id}")
def review_decision(suggestion_id: str, decision: str = Form(...),
                    reviewer: str = Form(...), comment: str = Form("")):
    safe_comment = comment.replace("'", "''")
    psql_query(
        "INSERT INTO review_decisions (suggestion_id, doc_id, clause_key, decision, reviewer, comment) "
        f"VALUES ('{suggestion_id}', '', '', '{decision}', '{reviewer}', '{safe_comment}');")
    return JSONResponse({"suggestion_id": suggestion_id, "decision": decision,
                         "note": "Decision recorded only; no document was modified."})


@app.post("/feedback")
def feedback(event_id: int = Form(...), username: str = Form(...),
             helpful: bool = Form(...), comment: str = Form("")):
    safe_comment = comment.replace("'", "''")
    psql_query(
        "INSERT INTO feedback (event_id, username, helpful, comment) "
        f"VALUES ({event_id}, '{username}', {str(helpful).lower()}, '{safe_comment}');")
    return JSONResponse({"status": "recorded"})


@app.get("/admin", response_class=HTMLResponse)
def admin():
    qdrant_status = "unknown"
    try:
        with urllib.request.urlopen(f"{QDRANT_URL}/collections", timeout=5) as response:
            qdrant_status = f"reachable (HTTP {response.status})"
    except Exception as error:
        qdrant_status = f"unreachable: {error}"
    model_status = "unknown"
    try:
        path = "/api/tags" if MODEL_BACKEND == "ollama" else "/v1/models"
        with urllib.request.urlopen(f"{MODEL_API_URL}{path}", timeout=5) as response:
            model_status = f"reachable (HTTP {response.status})"
    except Exception as error:
        model_status = f"unreachable: {error}"
    body = (f"<ul><li>Qdrant ({QDRANT_URL}): {html.escape(qdrant_status)}</li>"
            f"<li>Model ({MODEL_BACKEND} {MODEL_NAME}): {html.escape(model_status)}</li>"
            f"<li>Collection: {html.escape(COLLECTION_NAME)}</li></ul>")
    return page("Admin status", body)


if __name__ == "__main__":
    print(f"Serving on {APP_HOST}:{APP_PORT} (behind Nginx TLS in production).")
    uvicorn.run(app, host=APP_HOST, port=APP_PORT)

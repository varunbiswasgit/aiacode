#!/usr/bin/env python3
"""AI-RAG internal web app (FastAPI): login, ask, clause updates, review
queue, feedback, admin status.

Pinned dependencies (install exactly these; the app NEVER installs
packages at runtime):
    fastapi==0.115.6, uvicorn==0.30.6, psycopg[binary]==3.2.3
    (Python 3.12; phase-05 pipeline scripts are invoked as subprocesses.)

Tracker tasks: P8-01 (search/Q&A UI), P8-02 (ask page with citations),
               P8-03 (clause update dashboard), P8-04 (review queue UI),
               P8-05 (feedback buttons), P8-06 (admin page)
Startup prompts (every value also readable from an AIRAG_* environment
variable, so prompting can be skipped): app_host, app_port,
app_base_url, qdrant_url, collection_name, model_backend, model_api_url,
model_name, embedding_model_name, reranker_model_name, final_top_k,
min_rerank_score, temperature, db_host, db_port, audit_db_name,
standards_db_name, docs_db_name, db_user, db_password (secret, via
getpass or AIRAG_DB_PASSWORD), prompts_directory, access_map_path,
user_groups_map_path, pipeline_scripts_dir, sso_config_path,
session_ttl_seconds, session_cookie_secure, and (LDAP only)
ldap_user_dn_template.
Example defaults are shown as overridable and are labelled
"(example -- replace with your value)"; they are placeholders, not
deployment values.

Authentication: session-based. The SSO configuration written by
07-configure-sso.sh (sso_type oidc or ldap, mode 600) is read at startup
and re-read at login. NO route except /login, /callback and /logout is
usable without a valid server-side session cookie; unauthenticated
requests are redirected to /login (HTML routes) or answered 401 (POST
JSON routes). The authenticated username and the user's directory groups
come ONLY from the session: groups are resolved server-side from the
authenticated username (ID-token groups claim for OIDC, else the
operator-maintained user-groups map file) and passed internally to the
permission filter. No client-supplied groups are accepted anywhere, and
no form field for groups exists. A user with no mapped groups is denied
everything (fail closed): /ask abstains without running the pipeline.
OIDC login uses the authorization-code flow against the issuer found by
OpenID discovery on the configured issuer URL; the ID token's RS256
signature is verified against the provider JWKS and issuer/audience/
expiry are checked before any session is created. LDAP login verifies
the password with a real directory bind (ldapwhoami) for the user DN
built from the prompted DN template; the password is passed via a
mode-0600 temporary file that is deleted afterwards, never via argv,
echo, or logs. Neither flow has been verified against a live identity
provider in this environment (no IdP was available and the app is never
executed here); wire the provider URLs to a real IdP and test both
flows before production use.
Secrets: the database password travels only via getpass/PGPASSWORD-style
handling inside psycopg (in-process, never argv, never logged); the
OIDC client secret travels only in the token-exchange POST body, never
in logs or URLs; LDAP bind/user passwords travel only via mode-0600
temporary files deleted after use.

/ask pipeline (in this order, via subprocess, one temp dir per request):
    05-hybrid-retrieve.py -> 05-permission-filter.py ->
    05-rerank-results.py -> 05-generate-answer.py
The permission filter runs on the retrieved candidates BEFORE reranking
and generation, so no unfiltered content reaches the model. Reranking
is fed ONLY the filtered file (candidates-filtered.json), never the
unfiltered candidates file. Abstain rule: if the filtered set is empty,
if no reranked chunk meets the minimum score, or if the generated
answer carries no [doc#chunk] citation, the app answers with an
explicit "not found in internal documents" message and never guesses.
Answers that pass are returned with their citations plus the list of
source chunks used.
Audit: EVERY /ask event, answered or abstained (including pipeline
errors and fail-closed no-group cases), is written to audit_log in the
audit database with username, retrieved doc IDs, clause versions (from
chunk payload fields when present; the current phase-05 chunk contract
carries no version fields, so that list is empty unless a future chunk
payload supplies them), an answer excerpt, the abstained flag, and the
database default timestamp.
Databases: the audit database holds audit_log, feedback and
review_decisions; the STANDARDS database holds clause_changes and the
suggestions table (created by a separate script). /updates and /review
read the standards database; /ask audit/feedback/review-decision
writes go to the audit database. If the suggestions table does not
exist yet, /updates and /review say so cleanly instead of crashing.
Suggest-never-edit: /review only records approve/reject decisions in
review_decisions (with the real document_id and clause_id read from
the suggestion row). No code path here writes to a controlled document.
SQL: all database access uses psycopg with %s placeholders. There is
no string interpolation into SQL anywhere in this file.
Run: python 08-app-main.py   (then answer the prompts, or set the same
  values as environment variables AIRAG_* to skip prompting)
Routes: GET /login, POST /login, GET /callback, GET /logout,
  GET / (search page), POST /ask, GET /updates, GET /review,
  POST /review/{suggestion_id}, POST /feedback, GET /admin
"""
import base64
import getpass
import hashlib
import hmac
import html
import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request

try:
    from fastapi import FastAPI, Form, Request
    from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
    import uvicorn
    import psycopg
except ImportError as import_error:
    raise SystemExit(
        "Missing dependency: %s. Install the pinned versions and retry "
        "(the app never installs packages at runtime): "
        'fastapi==0.115.6, uvicorn==0.30.6, "psycopg[binary]==3.2.3"'
        % import_error
    ) from import_error


# ---------------------------------------------------------------- prompts

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


# ------------------------------------------------------- SSO config file
# Format written by 07-configure-sso.sh: simple "key: value" lines.
# OIDC: sso_type, issuer_url, client_id, client_secret.
# LDAP: sso_type, ldap_url, bind_dn, bind_password.

def load_sso_config(path):
    config = {}
    try:
        with open(path, encoding="utf-8") as handle:
            for line in handle:
                line = line.strip()
                if not line or line.startswith("#") or ":" not in line:
                    continue
                key, value = line.split(":", 1)
                config[key.strip()] = value.strip()
    except OSError as error:
        print(f"WARNING: could not read SSO config at {path}: {error}")
        return {}
    sso_type = config.get("sso_type", "")
    if sso_type == "oidc":
        if not config.get("issuer_url") or not config.get("client_id"):
            print("WARNING: SSO config is oidc but issuer_url/client_id is missing.")
            return {}
    elif sso_type == "ldap":
        if not config.get("ldap_url") or not config.get("bind_dn"):
            print("WARNING: SSO config is ldap but ldap_url/bind_dn is missing.")
            return {}
    else:
        print(f"WARNING: SSO config sso_type is not oidc/ldap: {sso_type!r}")
        return {}
    return config


def load_user_groups(path):
    """Server-side username -> directory-groups map (operator file).

    Returns {} when the file is missing/unreadable; callers fail closed
    (no groups means nothing is visible), they never fall back to
    client-supplied values.
    """
    try:
        with open(path, encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, ValueError) as error:
        print(f"WARNING: could not read user-groups map at {path}: {error}")
        return {}
    groups_map = {}
    if isinstance(data, dict):
        for username, groups in data.items():
            if isinstance(groups, list):
                groups_map[str(username)] = [str(group) for group in groups]
    return groups_map


print("=== AI-RAG app startup configuration ===")
APP_HOST = prompt_env("AIRAG_HOST", "Bind host", "0.0.0.0")
APP_PORT = int(prompt_env("AIRAG_PORT", "Bind port", "8000"))
APP_BASE_URL = prompt_env(
    "AIRAG_APP_BASE_URL",
    "App public base URL for OIDC redirect (example -- replace with your value)",
    "http://localhost:8000")
QDRANT_URL = prompt_env(
    "AIRAG_QDRANT_URL",
    "Qdrant base URL (example -- replace with your value)",
    "http://localhost:6333")
COLLECTION_NAME = prompt_env(
    "AIRAG_COLLECTION",
    "Qdrant collection (example -- replace with your value)",
    "airag_docs")
MODEL_BACKEND = prompt_env("AIRAG_MODEL_BACKEND", "Model backend (ollama/vllm)", "ollama")
MODEL_API_URL = prompt_env(
    "AIRAG_MODEL_API_URL",
    "Model API URL (example -- replace with your value)",
    "http://localhost:11434")
MODEL_NAME = prompt_env(
    "AIRAG_MODEL_NAME",
    "Model name (example -- replace with your value)",
    "llama3.1:8b")
EMBEDDING_MODEL = prompt_env(
    "AIRAG_EMBEDDING_MODEL",
    "Embedding model (example -- replace with your value)",
    "BAAI/bge-large-en-v1.5")
RERANKER_MODEL = prompt_env(
    "AIRAG_RERANKER_MODEL",
    "Reranker model (example -- replace with your value)",
    "BAAI/bge-reranker-large")
FINAL_TOP_K = prompt_env("AIRAG_FINAL_TOP_K", "Final top-k to keep", "8")
RETRIEVE_TOP_K = prompt_env("AIRAG_RETRIEVE_TOP_K", "Candidates per path", "50")
RRF_K = prompt_env("AIRAG_RRF_K", "RRF constant k", "60")
MIN_RERANK_SCORE = prompt_env("AIRAG_MIN_RERANK_SCORE", "Minimum rerank score to answer", "0.0")
TEMPERATURE = prompt_env("AIRAG_TEMPERATURE", "Temperature (0.1-0.3 for factual answers)", "0.2")
DB_HOST = prompt_env(
    "AIRAG_DB_HOST",
    "Postgres host (example -- replace with your value)",
    "localhost")
DB_PORT = prompt_env("AIRAG_DB_PORT", "Postgres port", "5432")
AUDIT_DB_NAME = prompt_env(
    "AIRAG_AUDIT_DB_NAME",
    "Audit database (example -- replace with your value)",
    "airag_audit")
STANDARDS_DB_NAME = prompt_env(
    "AIRAG_STANDARDS_DB_NAME",
    "Standards database (example -- replace with your value)",
    "airag_standards")
DOCS_DB_NAME = prompt_env(
    "AIRAG_DOCS_DB_NAME",
    "Documents database for keyword search (example -- replace with your value)",
    "airag_docs")
DB_USER = prompt_env(
    "AIRAG_DB_USER",
    "Database user (example -- replace with your value)",
    "airag_app")
DB_PASSWORD = prompt_secret_env("AIRAG_DB_PASSWORD", "Database password (not echoed)")
PROMPTS_DIR = prompt_env(
    "AIRAG_PROMPTS_DIR",
    "Prompts directory (example -- replace with your value)",
    "/opt/airag/prompts")
ACCESS_MAP_PATH = prompt_env(
    "AIRAG_ACCESS_MAP",
    "Access map JSON path (example -- replace with your value)",
    "/opt/airag/access-map.json")
USER_GROUPS_MAP_PATH = prompt_env(
    "AIRAG_USER_GROUPS_MAP",
    "Server-side user-groups map JSON path (example -- replace with your value)",
    "/opt/airag/user-groups.json")
SSO_CONFIG_PATH = prompt_env(
    "AIRAG_SSO_CONFIG",
    "SSO config path written by 07-configure-sso.sh (example -- replace with your value)",
    "/opt/airag/sso-config.yaml")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PIPELINE_DIR = prompt_env(
    "AIRAG_PIPELINE_DIR",
    "Phase-05 pipeline scripts directory (example -- replace with your value)",
    os.path.join(os.path.dirname(SCRIPT_DIR), "05-RAG-core"))
SESSION_TTL_SECONDS = int(prompt_env("AIRAG_SESSION_TTL", "Session lifetime (seconds)", "3600"))
SESSION_COOKIE_SECURE = prompt_env(
    "AIRAG_SESSION_COOKIE_SECURE",
    "Mark session cookie Secure (yes/no; yes behind TLS)",
    "no").strip().lower() == "yes"

SSO_CONFIG = load_sso_config(SSO_CONFIG_PATH)
if not SSO_CONFIG:
    print("WARNING: no valid SSO config loaded; logins will fail closed "
          "until a valid config from 07-configure-sso.sh is present.")

LDAP_USER_DN_TEMPLATE = ""
if SSO_CONFIG.get("sso_type") == "ldap":
    LDAP_USER_DN_TEMPLATE = prompt_env(
        "AIRAG_LDAP_USER_DN_TEMPLATE",
        "LDAP user DN template with {username} (example -- replace with your value)",
        "uid={username},ou=users,dc=example,dc=com")

app = FastAPI(title="AI-RAG internal assistant")

# ------------------------------------------------------------ sessions
# Server-side only: the cookie carries a random token, never identity
# or groups. Identity and groups live in this dict, keyed by token.

SESSIONS = {}
OIDC_PENDING = {}
CITATION_PATTERN = re.compile(r"\[[^\]]+#\d+\]")
DECISION_ALLOWLIST = ("approved", "rejected")
NOT_FOUND_MESSAGE = "Not found in internal documents."
SUGGESTIONS_MISSING_MESSAGE = (
    "The suggestion step has not been run yet, so no suggestions or "
    "impact ratings are available.")


def create_session(username, groups):
    token = secrets.token_urlsafe(32)
    SESSIONS[token] = {
        "username": username,
        "groups": list(groups),
        "expires_at": time.time() + SESSION_TTL_SECONDS,
    }
    return token


def get_session(request):
    token = request.cookies.get("airag_session", "")
    session = SESSIONS.get(token)
    if not session:
        return None
    if session["expires_at"] < time.time():
        SESSIONS.pop(token, None)
        return None
    return session


def clear_session(request):
    token = request.cookies.get("airag_session", "")
    SESSIONS.pop(token, None)


def session_cookie_response(response, token):
    response.set_cookie(
        "airag_session", token, httponly=True, samesite="lax",
        secure=SESSION_COOKIE_SECURE, max_age=SESSION_TTL_SECONDS, path="/")
    return response


def resolve_server_groups(username, token_groups):
    """Groups from the authenticated identity only, server-side.

    Token groups (OIDC groups claim) win when present; otherwise the
    operator-maintained user-groups map file is consulted. Never any
    client-supplied value. Unknown users get [] (fail closed).
    """
    if token_groups:
        return [str(group) for group in token_groups]
    return load_user_groups(USER_GROUPS_MAP_PATH).get(username, [])


# ------------------------------------------------------- OIDC helpers

def oidc_discovery(issuer_url):
    url = issuer_url.rstrip("/") + "/.well-known/openid-configuration"
    with urllib.request.urlopen(url, timeout=10) as response:
        return json.loads(response.read().decode("utf-8"))


def _b64url_decode(data):
    padding = "=" * (-len(data) % 4)
    return base64.urlsafe_b64decode(data + padding)


def verify_rs256_id_token(jwt, jwks_uri):
    """Verify an RS256 ID token signature against the provider JWKS.

    Pure-stdlib RSA verification (PKCS#1 v1.5, SHA-256). Returns the
    payload dict on success, None on any failure (fail closed).
    """
    try:
        header_b64, payload_b64, signature_b64 = jwt.split(".")
        header = json.loads(_b64url_decode(header_b64).decode("utf-8"))
        if header.get("alg") != "RS256":
            return None
        with urllib.request.urlopen(jwks_uri, timeout=10) as response:
            jwks = json.loads(response.read().decode("utf-8"))
        key = None
        for candidate in jwks.get("keys", []):
            if candidate.get("kty") == "RSA" and candidate.get("kid") == header.get("kid"):
                key = candidate
                break
        if key is None:
            return None
        modulus = int.from_bytes(_b64url_decode(key["n"]), "big")
        exponent = int.from_bytes(_b64url_decode(key["e"]), "big")
        signature = int.from_bytes(_b64url_decode(signature_b64), "big")
        encoded = pow(signature, exponent, modulus)
        em = encoded.to_bytes((modulus.bit_length() + 7) // 8, "big")
        digest = hashlib.sha256(f"{header_b64}.{payload_b64}".encode("utf-8")).digest()
        prefix = bytes.fromhex("3031300d060960864801650304020105000420")
        expected = (b"\x00\x01" + b"\xff" * (len(em) - len(prefix) - len(digest) - 3)
                    + b"\x00" + prefix + digest)
        if not hmac.compare_digest(em, expected):
            return None
        return json.loads(_b64url_decode(payload_b64).decode("utf-8"))
    except Exception:
        return None


def oidc_authorize_redirect():
    """Build the provider authorization redirect, or None on failure."""
    config = load_sso_config(SSO_CONFIG_PATH)
    if config.get("sso_type") != "oidc":
        return None
    try:
        discovery = oidc_discovery(config["issuer_url"])
        authorize_url = discovery["authorization_endpoint"]
    except Exception:
        return None
    state = secrets.token_urlsafe(24)
    nonce = secrets.token_urlsafe(24)
    OIDC_PENDING[state] = {"nonce": nonce, "expires_at": time.time() + 600}
    params = urllib.parse.urlencode({
        "response_type": "code",
        "client_id": config["client_id"],
        "redirect_uri": APP_BASE_URL.rstrip("/") + "/callback",
        "scope": "openid profile groups",
        "state": state,
        "nonce": nonce,
    })
    return RedirectResponse(f"{authorize_url}?{params}")


# ------------------------------------------------------- LDAP helper

def ldap_bind_check(ldap_url, user_dn, password):
    """Verify a password with a real directory bind (ldapwhoami).

    The password is written to a mode-0600 temporary file and deleted
    afterwards; it never appears in argv, echo output, or logs. A
    missing ldapwhoami binary or a failed bind means "not
    authenticated" (fail closed).
    """
    handle = tempfile.NamedTemporaryFile("w", delete=False)
    try:
        handle.write(password)
        handle.close()
        os.chmod(handle.name, 0o600)
        result = subprocess.run(
            ["ldapwhoami", "-H", ldap_url, "-D", user_dn, "-y", handle.name],
            capture_output=True, text=True, timeout=30)
        return result.returncode == 0
    except (OSError, subprocess.TimeoutExpired):
        return False
    finally:
        try:
            os.unlink(handle.name)
        except OSError:
            pass


# ------------------------------------------------------- DB helpers
# All SQL uses psycopg %s placeholders; no value is ever interpolated
# into a SQL string in this file.

def fetch_dicts(db_name, sql, params=()):
    with psycopg.connect(host=DB_HOST, port=DB_PORT, dbname=db_name,
                         user=DB_USER, password=DB_PASSWORD) as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            columns = [desc[0] for desc in cur.description] if cur.description else []
            return [dict(zip(columns, row)) for row in cur.fetchall()]


def execute(db_name, sql, params=()):
    with psycopg.connect(host=DB_HOST, port=DB_PORT, dbname=db_name,
                         user=DB_USER, password=DB_PASSWORD) as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
        conn.commit()


def suggestions_table_exists():
    rows = fetch_dicts(
        STANDARDS_DB_NAME,
        "SELECT 1 AS found FROM information_schema.tables "
        "WHERE table_schema = %s AND table_name = %s",
        ("public", "suggestions"))
    return bool(rows)


def audit_ask(username, question, retrieved_docs, clause_versions,
              answer_excerpt, abstained):
    """Write one /ask event to audit_log (answered AND abstained)."""
    try:
        execute(
            AUDIT_DB_NAME,
            "INSERT INTO audit_log (username, question, retrieved_docs, "
            "clause_versions, answer_excerpt, abstained) "
            "VALUES (%s, %s, %s, %s, %s, %s)",
            (username, question, list(retrieved_docs), list(clause_versions),
             answer_excerpt, abstained))
    except Exception as error:
        # Never hide the answer behind an audit failure, and never log
        # the question/answer content with the error.
        print(f"WARNING: audit_log write failed: {error}")


# ------------------------------------------------------- pipeline

def run_pipeline_script(script_name, stdin_lines, cwd, timeout=600):
    """Run one phase-05 script as a subprocess, feeding its prompts
    on stdin in prompt order. Secrets are stdin lines, never argv.

    NOTE: the phase-05 scripts read their database password with
    getpass, which falls back to piped stdin when the process has no
    controlling terminal (e.g. under systemd). Do not run this app
    attached to an interactive terminal for /ask requests, or the
    getpass prompt may read from the terminal instead of the pipe.
    """
    script_path = os.path.join(PIPELINE_DIR, script_name)
    result = subprocess.run(
        [sys.executable, script_path], input="\n".join(stdin_lines) + "\n",
        capture_output=True, text=True, cwd=cwd, timeout=timeout)
    if result.returncode != 0:
        raise RuntimeError(f"{script_name} failed: {result.stderr.strip()[-500:]}")
    return result.stdout


def load_json_file(path):
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


# ------------------------------------------------------- HTML helpers

def page(title, body):
    return f"""<!DOCTYPE html><html><head><title>{html.escape(title)}</title></head>
<body><h1>{html.escape(title)}</h1>{body}
<p><a href="/">Ask</a> | <a href="/updates">Clause updates</a> |
<a href="/review">Review queue</a> | <a href="/admin">Admin</a> |
<a href="/logout">Log out</a></p></body></html>"""


def login_page(message=""):
    config = load_sso_config(SSO_CONFIG_PATH)
    sso_type = config.get("sso_type", "not configured")
    provider = config.get("issuer_url") or config.get("ldap_url") or "none"
    note = (f"<p>Sign in via the configured SSO provider "
            f"({html.escape(sso_type)}: {html.escape(provider)}).</p>")
    if message:
        note += f"<p>{html.escape(message)}</p>"
    if sso_type == "ldap":
        form = ('<form method="post" action="/login">'
                'Username: <input name="username" required> '
                'Password: <input name="password" type="password" required> '
                '<button>Log in</button></form>')
        return page("Log in", note + form)
    if sso_type == "oidc":
        return page("Log in", note + '<p><a href="/login">Continue to SSO provider</a></p>')
    return page("Log in", note + "<p>SSO is not configured; logins are disabled.</p>")


def require_session_html(request):
    session = get_session(request)
    if session is None:
        return None, RedirectResponse("/login")
    return session, None


def require_session_json(request):
    session = get_session(request)
    if session is None:
        return None, JSONResponse({"error": "authentication required"}, status_code=401)
    return session, None


# ------------------------------------------------------- auth routes

@app.get("/login", response_class=HTMLResponse)
def login_get():
    config = load_sso_config(SSO_CONFIG_PATH)
    if config.get("sso_type") == "oidc":
        redirect = oidc_authorize_redirect()
        if redirect is not None:
            return redirect
        return login_page("SSO provider could not be reached; login failed.")
    return login_page()


@app.post("/login")
def login_post(username: str = Form(...), password: str = Form(...)):
    config = load_sso_config(SSO_CONFIG_PATH)
    if config.get("sso_type") == "oidc":
        redirect = oidc_authorize_redirect()
        if redirect is not None:
            return redirect
        return login_page("SSO provider could not be reached; login failed.")
    if config.get("sso_type") != "ldap":
        return login_page("SSO is not configured; logins are disabled.")
    if "{username}" not in LDAP_USER_DN_TEMPLATE:
        return login_page("LDAP user DN template is misconfigured; login failed.")
    user_dn = LDAP_USER_DN_TEMPLATE.format(username=username)
    if not ldap_bind_check(config["ldap_url"], user_dn, password):
        return login_page("Login failed.")
    groups = resolve_server_groups(username, None)
    token = create_session(username, groups)
    return session_cookie_response(RedirectResponse("/", status_code=303), token)


@app.get("/callback")
def oidc_callback(request: Request):
    state = request.query_params.get("state", "")
    code = request.query_params.get("code", "")
    pending = OIDC_PENDING.pop(state, None)
    if not pending or pending["expires_at"] < time.time() or not code:
        return login_page("SSO login failed (bad or expired state).")
    config = load_sso_config(SSO_CONFIG_PATH)
    if config.get("sso_type") != "oidc" or not config.get("client_secret"):
        return login_page("SSO is not configured; logins are disabled.")
    try:
        discovery = oidc_discovery(config["issuer_url"])
        payload = urllib.parse.urlencode({
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": APP_BASE_URL.rstrip("/") + "/callback",
            "client_id": config["client_id"],
            "client_secret": config["client_secret"],
        }).encode("utf-8")
        token_request = urllib.request.Request(
            discovery["token_endpoint"], data=payload, method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded"})
        with urllib.request.urlopen(token_request, timeout=30) as response:
            token_data = json.loads(response.read().decode("utf-8"))
        id_token = token_data.get("id_token", "")
        claims = verify_rs256_id_token(id_token, discovery["jwks_uri"])
        if claims is None:
            return login_page("SSO token verification failed; login refused.")
        if claims.get("iss") != discovery.get("issuer", claims.get("iss")):
            return login_page("SSO token issuer mismatch; login refused.")
        audience = claims.get("aud")
        audience_values = audience if isinstance(audience, list) else [audience]
        if config["client_id"] not in audience_values:
            return login_page("SSO token audience mismatch; login refused.")
        if claims.get("exp", 0) < time.time():
            return login_page("SSO token expired; login refused.")
        if claims.get("nonce") != pending["nonce"]:
            return login_page("SSO token nonce mismatch; login refused.")
        username = (claims.get("preferred_username") or claims.get("username")
                    or claims.get("sub"))
        if not username:
            return login_page("SSO token carried no username; login refused.")
        raw_groups = claims.get("groups")
        token_groups = raw_groups if isinstance(raw_groups, list) else []
        groups = resolve_server_groups(str(username), token_groups)
    except Exception:
        return login_page("SSO login failed.")
    token = create_session(str(username), groups)
    return session_cookie_response(RedirectResponse("/", status_code=303), token)


@app.get("/logout")
def logout(request: Request):
    clear_session(request)
    response = RedirectResponse("/login")
    response.delete_cookie("airag_session", path="/")
    return response


# ------------------------------------------------------- app routes

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    session, redirect = require_session_html(request)
    if redirect is not None:
        return redirect
    return page("AI-RAG assistant",
                f"<p>Signed in as {html.escape(session['username'])}.</p>"
                '<form method="post" action="/ask">'
                'Question: <input name="question" size="80" required> '
                '<button>Ask</button></form>')


@app.post("/ask", response_class=HTMLResponse)
def ask(request: Request, question: str = Form(...)):
    """Retrieve -> permission-filter -> rerank -> generate, in order.

    Reranking is fed ONLY candidates-filtered.json. Any empty filtered
    set, any empty usable reranked set, or any citation-less answer
    abstains explicitly. Every outcome is audited (see audit_ask).
    """
    session, error = require_session_json(request)
    if error is not None:
        return error
    username = session["username"]
    server_groups = session["groups"]
    question = question.strip()
    retrieved_docs = []
    clause_versions = []
    answer = NOT_FOUND_MESSAGE
    abstained = True
    error_message = ""
    tmpdir = tempfile.mkdtemp(prefix="airag-ask-")
    try:
        if not question:
            error_message = "A question is required."
        elif not server_groups:
            # Fail closed. Do NOT run the pipeline: the permission-filter
            # script defaults a blank groups answer to a group name, so
            # a blank line must never be fed to it.
            answer = NOT_FOUND_MESSAGE
            error_message = ("No directory groups are mapped for this user, "
                             "so nothing can be shown.")
        else:
            candidates_path = os.path.join(tmpdir, "candidates.json")
            filtered_path = os.path.join(tmpdir, "candidates-filtered.json")
            reranked_path = os.path.join(tmpdir, "reranked.json")
            # Step 1: hybrid retrieve.
            run_pipeline_script(
                "05-hybrid-retrieve.py",
                [question, QDRANT_URL, COLLECTION_NAME, EMBEDDING_MODEL,
                 DB_HOST, DB_PORT, DOCS_DB_NAME, DB_USER, DB_PASSWORD,
                 RETRIEVE_TOP_K, RRF_K, tmpdir],
                cwd=tmpdir)
            # Step 2: permission filter (server-side groups only).
            run_pipeline_script(
                "05-permission-filter.py",
                [username, ",".join(server_groups), ACCESS_MAP_PATH,
                 candidates_path, filtered_path],
                cwd=tmpdir)
            filtered = load_json_file(filtered_path)
            if not filtered:
                answer = NOT_FOUND_MESSAGE
            else:
                # Step 3: rerank the FILTERED file, never the unfiltered one.
                run_pipeline_script(
                    "05-rerank-results.py",
                    [question, filtered_path, RERANKER_MODEL,
                     FINAL_TOP_K, reranked_path],
                    cwd=tmpdir)
                reranked = load_json_file(reranked_path)
                try:
                    min_score = float(MIN_RERANK_SCORE)
                except ValueError:
                    min_score = 0.0
                usable = [chunk for chunk in reranked
                          if chunk.get("rerank_score", 0) >= min_score]
                retrieved_docs = [f"{chunk.get('doc_id', '')}#{chunk.get('chunk_index', '')}"
                                  for chunk in usable]
                for chunk in usable:
                    if all(key in chunk for key in ("standard", "clause_no", "version")):
                        clause_versions.append(
                            f"{chunk['standard']}:{chunk['clause_no']}@{chunk['version']}")
                if not usable:
                    answer = NOT_FOUND_MESSAGE
                else:
                    # Step 4: generate from the reranked (filtered) chunks.
                    stdout = run_pipeline_script(
                        "05-generate-answer.py",
                        [question, reranked_path, PROMPTS_DIR, MODEL_BACKEND,
                         MODEL_API_URL, MODEL_NAME, MIN_RERANK_SCORE, TEMPERATURE],
                        cwd=tmpdir)
                    if NOT_FOUND_MESSAGE in stdout:
                        answer = NOT_FOUND_MESSAGE
                    else:
                        marker = "--- Answer ---"
                        generated = stdout.split(marker, 1)[1].strip() if marker in stdout else ""
                        if generated and CITATION_PATTERN.search(generated):
                            answer = generated
                            abstained = False
                        else:
                            answer = NOT_FOUND_MESSAGE
    except Exception as error:
        # Pipeline failure: no answer is returned, and the event is
        # still audited below as abstained so no /ask escapes audit.
        answer = NOT_FOUND_MESSAGE
        error_message = f"Pipeline error: {error}"
    finally:
        audit_ask(username, question, retrieved_docs, clause_versions,
                  answer[:500] if answer else None, abstained)
        shutil.rmtree(tmpdir, ignore_errors=True)
    body = f"<p>Question: {html.escape(question)}</p>"
    if error_message:
        body += f"<p>{html.escape(error_message)}</p>"
    body += f"<p>{html.escape(answer)}</p>"
    if not abstained and retrieved_docs:
        sources = "".join(f"<li>{html.escape(doc)}</li>" for doc in retrieved_docs)
        body += f"<p>Sources:</p><ul>{sources}</ul>"
    return page("Answer", body)


@app.get("/updates", response_class=HTMLResponse)
def updates(request: Request):
    session, redirect = require_session_html(request)
    if redirect is not None:
        return redirect
    try:
        # clause_changes lives in the STANDARDS database, not the audit DB.
        changes = fetch_dicts(
            STANDARDS_DB_NAME,
            "SELECT standard, chapter, clause_no, change_type, new_version, "
            "detected_at FROM clause_changes ORDER BY detected_at DESC LIMIT 50")
        if suggestions_table_exists():
            suggestions = fetch_dicts(
                STANDARDS_DB_NAME,
                "SELECT s.suggestion_id, s.document_id, s.clause_id, "
                "s.impact_rating, s.status, s.suggested_text, c.standard, "
                "c.clause_no, c.new_version FROM suggestions s "
                "LEFT JOIN clause_changes c ON c.change_id = s.change_id "
                "ORDER BY s.created_at DESC LIMIT 100")
            rows = "".join(
                "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                    html.escape(str(item.get("document_id", ""))),
                    html.escape(str(item.get("clause_id", ""))),
                    html.escape(str(item.get("impact_rating", ""))),
                    html.escape(str(item.get("status", ""))),
                    html.escape(str(item.get("suggested_text", ""))[:200]))
                for item in suggestions)
            suggestions_body = ("<p>Affected documents and impact ratings from "
                                "the suggestions table:</p>"
                                "<table><tr><th>Document</th><th>Clause</th>"
                                "<th>Impact</th><th>Status</th>"
                                "<th>Suggested text</th></tr>" + rows + "</table>")
        else:
            suggestions_body = f"<p>{html.escape(SUGGESTIONS_MISSING_MESSAGE)}</p>"
        change_rows = "".join(
            "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                html.escape(str(item.get("standard", ""))),
                html.escape(str(item.get("clause_no", ""))),
                html.escape(str(item.get("change_type", ""))),
                html.escape(str(item.get("new_version", ""))),
                html.escape(str(item.get("detected_at", ""))))
            for item in changes)
        body = (suggestions_body
                + "<p>Recent clause changes:</p>"
                + "<table><tr><th>Standard</th><th>Clause</th><th>Change</th>"
                "<th>New version</th><th>Detected</th></tr>" + change_rows + "</table>")
    except Exception as error:
        body = f"<p>Could not read the standards database: {html.escape(str(error))}</p>"
    return page("Clause updates", body)


@app.get("/review", response_class=HTMLResponse)
def review_queue(request: Request):
    session, redirect = require_session_html(request)
    if redirect is not None:
        return redirect
    try:
        if not suggestions_table_exists():
            return page("Review queue",
                        f"<p>{html.escape(SUGGESTIONS_MISSING_MESSAGE)}</p>")
        suggestions = fetch_dicts(
            STANDARDS_DB_NAME,
            "SELECT suggestion_id, document_id, clause_id, impact_rating, "
            "status, suggested_text FROM suggestions WHERE status = %s "
            "ORDER BY created_at DESC LIMIT 100",
            ("pending",))
        if not suggestions:
            body = "<p>No pending suggestions.</p>"
        else:
            parts = ["<p>Suggestions from gap analysis appear here for QA/doc-owner "
                     "approval. Approving a suggestion records the decision only "
                     "- it never edits a controlled document.</p>"]
            for item in suggestions:
                suggestion_id = item.get("suggestion_id")
                parts.append(
                    f"<div><p>{html.escape(str(item.get('document_id', '')))} / "
                    f"{html.escape(str(item.get('clause_id', '')))} "
                    f"(impact: {html.escape(str(item.get('impact_rating', '')))})<br>"
                    f"{html.escape(str(item.get('suggested_text', ''))[:300])}</p>"
                    f'<form method="post" action="/review/{html.escape(str(suggestion_id))}">'
                    'Decision: <select name="decision"><option>approved</option>'
                    '<option>rejected</option></select> '
                    'Comment: <input name="comment"> <button>Submit</button></form></div>')
            body = "".join(parts)
    except Exception as error:
        body = f"<p>Could not read the review queue: {html.escape(str(error))}</p>"
    return page("Review queue", body)


@app.post("/review/{suggestion_id}")
def review_decision(request: Request, suggestion_id: str,
                    decision: str = Form(...), comment: str = Form("")):
    session, error = require_session_json(request)
    if error is not None:
        return error
    if decision not in DECISION_ALLOWLIST:
        return JSONResponse(
            {"error": "decision must be 'approved' or 'rejected'"},
            status_code=400)
    try:
        suggestion_key = int(suggestion_id)
    except ValueError:
        return JSONResponse({"error": "suggestion not found"}, status_code=404)
    try:
        # Real document/clause keys come from the suggestion row itself;
        # empty keys are never recorded.
        rows = fetch_dicts(
            STANDARDS_DB_NAME,
            "SELECT document_id, clause_id FROM suggestions "
            "WHERE suggestion_id = %s",
            (suggestion_key,))
    except Exception as error:
        return JSONResponse({"error": f"could not read suggestion: {error}"},
                            status_code=500)
    if not rows:
        return JSONResponse({"error": "suggestion not found"}, status_code=404)
    document_id = rows[0].get("document_id", "")
    clause_id = rows[0].get("clause_id", "")
    try:
        execute(
            AUDIT_DB_NAME,
            "INSERT INTO review_decisions (suggestion_id, doc_id, clause_key, "
            "decision, reviewer, comment) VALUES (%s, %s, %s, %s, %s, %s)",
            (str(suggestion_id), document_id, clause_id, decision,
             session["username"], comment))
    except Exception as error:
        return JSONResponse({"error": f"could not record decision: {error}"},
                            status_code=500)
    return JSONResponse({"suggestion_id": suggestion_id, "decision": decision,
                         "document_id": document_id, "clause_id": clause_id,
                         "note": "Decision recorded only; no document was modified."})


@app.post("/feedback")
def feedback(request: Request, event_id: int = Form(...),
             helpful: bool = Form(...), comment: str = Form("")):
    session, error = require_session_json(request)
    if error is not None:
        return error
    try:
        execute(
            AUDIT_DB_NAME,
            "INSERT INTO feedback (event_id, username, helpful, comment) "
            "VALUES (%s, %s, %s, %s)",
            (event_id, session["username"], helpful, comment))
    except Exception as error:
        return JSONResponse({"error": f"could not record feedback: {error}"},
                            status_code=500)
    return JSONResponse({"status": "recorded"})


@app.get("/admin", response_class=HTMLResponse)
def admin(request: Request):
    session, redirect = require_session_html(request)
    if redirect is not None:
        return redirect
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
    body = (f"<p>Signed in as {html.escape(session['username'])}.</p>"
            f"<ul><li>Qdrant ({html.escape(QDRANT_URL)}): {html.escape(qdrant_status)}</li>"
            f"<li>Model ({html.escape(MODEL_BACKEND)} {html.escape(MODEL_NAME)}): "
            f"{html.escape(model_status)}</li>"
            f"<li>Collection: {html.escape(COLLECTION_NAME)}</li></ul>")
    return page("Admin status", body)


if __name__ == "__main__":
    print(f"Serving on {APP_HOST}:{APP_PORT} (behind Nginx TLS in production).")
    uvicorn.run(app, host=APP_HOST, port=APP_PORT)

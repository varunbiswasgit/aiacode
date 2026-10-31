-- =============================================================================
-- Purpose:      Create the AI-RAG databases and the scoped application user.
-- Tracker task: P2-06
-- Run:         read -s -p 'Password for the application user: ' app_password
--              echo
--              sudo -u postgres psql -v docs_db=airag_docs \
--                   -v standards_db=airag_standards -v audit_db=airag_audit \
--                   -v app_user=airag_app -v app_password="$app_password" \
--                   -f 02-create-databases.sql
--              unset app_password
--              (The password is read silently by the shell, passed via -v,
--               and never written into this file or any log.)
-- Prerequisites: PostgreSQL 16 + pgvector installed. Run ON the DB node.
-- =============================================================================
\if :{?docs_db} \else \echo 'ERROR: missing -v docs_db=<name>' \quit \endif
\if :{?standards_db} \else \echo 'ERROR: missing -v standards_db=<name>' \quit \endif
\if :{?audit_db} \else \echo 'ERROR: missing -v audit_db=<name>' \quit \endif
\if :{?app_user} \else \echo 'ERROR: missing -v app_user=<name>' \quit \endif
\if :{?app_password} \else \echo 'ERROR: missing -v app_password=<password>' \quit \endif

-- Databases (idempotent: skip if they already exist)
SELECT 'CREATE DATABASE ' || :'docs_db' WHERE NOT EXISTS (SELECT FROM pg_database WHERE datname = :'docs_db')\gexec
SELECT 'CREATE DATABASE ' || :'standards_db' WHERE NOT EXISTS (SELECT FROM pg_database WHERE datname = :'standards_db')\gexec
SELECT 'CREATE DATABASE ' || :'audit_db' WHERE NOT EXISTS (SELECT FROM pg_database WHERE datname = :'audit_db')\gexec

-- Application user (created only if missing).
DO $$
BEGIN
    IF NOT EXISTS (SELECT FROM pg_roles WHERE rolname = :'app_user') THEN
        EXECUTE format('CREATE USER %I WITH PASSWORD %L', :'app_user', :'app_password');
    END IF;
END
$$;

-- Scoped grants: app user may connect to its own databases, nothing else.
GRANT CONNECT ON DATABASE :"docs_db" TO :"app_user";
GRANT CONNECT ON DATABASE :"standards_db" TO :"app_user";
GRANT CONNECT ON DATABASE :"audit_db" TO :"app_user";
\c :"docs_db"
CREATE EXTENSION IF NOT EXISTS vector;
GRANT USAGE, CREATE ON SCHEMA public TO :"app_user";
\c :"standards_db"
CREATE EXTENSION IF NOT EXISTS vector;
GRANT USAGE, CREATE ON SCHEMA public TO :"app_user";
\c :"audit_db"
GRANT USAGE, CREATE ON SCHEMA public TO :"app_user";

\echo 'Done. Databases and application user created. Store the password in your vault.'

-- =============================================================================
-- Purpose:      Create the AI-RAG databases and the scoped application user.
-- Tracker task: P2-06
-- Run:         read -s -p 'Password for the application user: ' PGPASSWORD
--              echo
--              export PGPASSWORD
--              sudo --preserve-env=PGPASSWORD -u postgres psql \
--                   -v docs_db=airag_docs \
--                   -v standards_db=airag_standards -v audit_db=airag_audit \
--                   -v app_user=airag_app \
--                   -f 02-create-databases.sql
--              unset PGPASSWORD
--              (All -v values above are examples — replace with your value.
--               The password is read silently by the shell with read -s
--               into the PGPASSWORD environment variable; inside this
--               script psql picks it up with \getenv (available since
--               PostgreSQL 15, so covered by the PostgreSQL 16
--               prerequisite) and it is never written into this file or
--               any log. Connection authentication here is the
--               operating-system peer login via sudo -u postgres, so
--               PGPASSWORD is not consumed as the connection password;
--               --preserve-env passes only the variable NAME on the sudo
--               command line, never the value. If sudo strips the
--               variable, this script stops with the missing-PGPASSWORD
--               error below instead of creating a user with no password.
--               FORBIDDEN: never pass the password as psql -v
--               app_password=... — command-line arguments are visible in
--               the process list (ps) to other users on the node.)
-- Prerequisites: PostgreSQL 16 + pgvector installed. Run ON the DB node.
-- Design note:  The app user gets USAGE on each schema only — never
--               CREATE — so it cannot create arbitrary tables. Tables are
--               created by the phase scripts running as the table owner
--               or a superuser (e.g. 04-create-clause-registry.sql and
--               07-create-audit-log.sql); table-level grants are issued
--               by 02-grant-database-permissions.sh (P2-15).
-- =============================================================================
\if :{?docs_db} \else \echo 'ERROR: missing -v docs_db=<name>' \quit \endif
\if :{?standards_db} \else \echo 'ERROR: missing -v standards_db=<name>' \quit \endif
\if :{?audit_db} \else \echo 'ERROR: missing -v audit_db=<name>' \quit \endif
\if :{?app_user} \else \echo 'ERROR: missing -v app_user=<name>' \quit \endif
\getenv app_password PGPASSWORD
\if :{?app_password} \else \echo 'ERROR: PGPASSWORD is not set. Read the application-user password silently into PGPASSWORD (read -s) before running psql; never pass it via -v app_password=... (visible in the process list).' \quit \endif

-- Databases (idempotent: skip if they already exist)
SELECT 'CREATE DATABASE ' || :'docs_db' WHERE NOT EXISTS (SELECT FROM pg_database WHERE datname = :'docs_db')\gexec
SELECT 'CREATE DATABASE ' || :'standards_db' WHERE NOT EXISTS (SELECT FROM pg_database WHERE datname = :'standards_db')\gexec
SELECT 'CREATE DATABASE ' || :'audit_db' WHERE NOT EXISTS (SELECT FROM pg_database WHERE datname = :'audit_db')\gexec

-- Application user (created only if missing).
DO $$
BEGIN
    IF :'app_password' = '' THEN
        RAISE EXCEPTION 'app_password must not be empty (PGPASSWORD was empty)';
    END IF;
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
GRANT USAGE ON SCHEMA public TO :"app_user";
\c :"standards_db"
CREATE EXTENSION IF NOT EXISTS vector;
GRANT USAGE ON SCHEMA public TO :"app_user";
\c :"audit_db"
GRANT USAGE ON SCHEMA public TO :"app_user";

\unset app_password
\echo 'Done. Databases and application user created. Store the password in your vault.'

-- =============================================================================
-- Purpose:      Create the clause registry tables in the standards database.
-- Tracker task: P4-07
-- Run:         sudo -u postgres psql -v standards_db=airag_standards \
--                   -v app_user=airag_app \
--                   -f 04-create-clause-registry.sql
--              (All -v values above are examples — replace with your value.
--               Run after 02-create-databases.sql so the app user already
--               exists; run as the table owner or a superuser so the
--               GRANT statements below succeed.)
-- Prerequisites: PostgreSQL 16 + pgvector (02 scripts). Run ON the DB node.
-- Design note: clauses are append-only (versions accumulate, never deleted).
--               The application user gets INSERT/SELECT only - no UPDATE/DELETE
--               - which enforces the version-retention rule (P4-11).
-- =============================================================================
\if :{?standards_db} \else \echo 'ERROR: missing -v standards_db=<name>' \quit \endif
\if :{?app_user} \else \echo 'ERROR: missing -v app_user=<name>' \quit \endif
\c :"standards_db"

CREATE EXTENSION IF NOT EXISTS vector;

CREATE TABLE IF NOT EXISTS clauses (
    clause_id     BIGSERIAL PRIMARY KEY,
    standard      TEXT NOT NULL,          -- 'FDA eCFR', 'EU GMP', 'ICH'
    chapter       TEXT NOT NULL,          -- e.g. 'Title 21 Part 211', 'Annex 11'
    clause_no     TEXT NOT NULL,          -- e.g. '211.100', '5.2'
    text          TEXT NOT NULL,
    version       DATE NOT NULL,          -- fetch date of this version
    source_url    TEXT NOT NULL DEFAULT '',
    embedding     VECTOR(1024),           -- set by 05-embed-documents.py; adjust dim to your model
    created_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (standard, chapter, clause_no, version)
);

CREATE TABLE IF NOT EXISTS clause_changes (
    change_id     BIGSERIAL PRIMARY KEY,
    standard      TEXT NOT NULL,
    chapter       TEXT NOT NULL,
    clause_no     TEXT NOT NULL,
    old_version   DATE,
    new_version   DATE NOT NULL,
    change_type   TEXT NOT NULL,          -- 'added', 'removed', 'modified'
    old_text      TEXT,
    new_text      TEXT,
    detected_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_clauses_lookup
    ON clauses (standard, chapter, clause_no, version DESC);
CREATE INDEX IF NOT EXISTS idx_changes_detected
    ON clause_changes (detected_at DESC);

-- Table-level grants (actual statements, INSERT/SELECT only).
GRANT INSERT, SELECT ON clauses TO :"app_user";
GRANT INSERT, SELECT ON clause_changes TO :"app_user";

\echo 'Done. Clause registry ready. The app user has INSERT/SELECT only — no UPDATE/DELETE on these tables.'

-- =============================================================================
-- Purpose:      Create the audit log tables (queries, answers, feedback).
-- Tracker task: P7-03
-- Run:         sudo -u postgres psql -v audit_db=airag_audit \
--                   -f 07-create-audit-log.sql
-- Prerequisites: PostgreSQL 16. Run ON the DB node.
-- Design note: audit tables are append-only. The app user gets INSERT and
--               SELECT only - never UPDATE or DELETE.
-- =============================================================================
\if :{?audit_db} \else \echo 'ERROR: missing -v audit_db=<name>' \quit \endif
\c :"audit_db"

CREATE TABLE IF NOT EXISTS audit_log (
    event_id        BIGSERIAL PRIMARY KEY,
    event_time      TIMESTAMPTZ NOT NULL DEFAULT now(),
    username        TEXT NOT NULL,
    question        TEXT NOT NULL,
    retrieved_docs  TEXT[] NOT NULL DEFAULT '{}',  -- doc_id#chunk_index values
    clause_versions TEXT[] NOT NULL DEFAULT '{}',  -- standard:clause_no@version
    answer_excerpt  TEXT,
    abstained       BOOLEAN NOT NULL DEFAULT FALSE
);

CREATE TABLE IF NOT EXISTS feedback (
    feedback_id     BIGSERIAL PRIMARY KEY,
    event_id        BIGINT NOT NULL REFERENCES audit_log (event_id),
    username        TEXT NOT NULL,
    helpful         BOOLEAN NOT NULL,
    comment         TEXT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS review_decisions (
    decision_id     BIGSERIAL PRIMARY KEY,
    suggestion_id   TEXT NOT NULL,
    doc_id          TEXT NOT NULL,
    clause_key      TEXT NOT NULL,
    decision        TEXT NOT NULL,   -- 'approved' or 'rejected'
    reviewer        TEXT NOT NULL,
    comment         TEXT,
    decided_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_audit_user_time ON audit_log (username, event_time DESC);
CREATE INDEX IF NOT EXISTS idx_feedback_event ON feedback (event_id);

\echo 'Done. Audit tables created. Grant the app user INSERT/SELECT only:'
\echo '  GRANT INSERT, SELECT ON audit_log, feedback, review_decisions TO <app_user>;'

AI-RAG Internal Systems is an on-premises retrieval-augmented generation platform for internal content discovery, question answering over internal documents and regulatory standards (FDA eCFR Title 21, EU GMP, ICH), and document-update recommendations when a standard changes. It is generic: every deployment value is supplied as user input at runtime, and no company-specific value is hardcoded. Every recommendation is a suggestion for human QA and document-owner approval; the system never auto-edits a controlled document.

# 09 — Eval and pilot

## Purpose

Prove the system works before go-live and learn from real users: build
the golden Q&A set, score answers against it, create the pilot AD
group, and report weekly pilot feedback. Fixing failure modes (P9-03),
pilot training (P9-05) and the go/no-go meeting (P9-07) stay manual.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 09-create-golden-set.py | P9-01 | Interactively builds golden.json: question, expected answer, expected source doc IDs, appended one item at a time. |
| 09-run-evaluation.py | P9-02 | Scores an answers JSONL against the golden set: retrieval hit rate, citation accuracy, abstain rate, keyword overlap; writes a Markdown report with the failure list. |
| 09-create-pilot-group.sh | P9-04 | Creates the pilot AD group and adds members (samba-tool, or prints the manual ADUC steps when samba-tool is absent). |
| 09-pilot-feedback-report.py | P9-06 | Queries the feedback table for totals, helpful rate, per-week breakdown and recent comments; writes a Markdown report. The temporary SQL file is deleted in try/finally, including on failure. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 09-create-golden-set.py | output_golden_path, then per item: question, expected_answer, expected_doc_ids |
| 09-run-evaluation.py | golden_json_path, answers_jsonl_path, output_report_path |
| 09-create-pilot-group.sh | ad_domain_name, pilot_group_name, group_member_usernames, ad_admin_username |
| 09-pilot-feedback-report.py | db_host, db_port, db_name, db_user, db_password (secret), report_weeks, output_report_path |

The database password is read with getpass and never saved. The answers
JSONL for evaluation is produced by running the app (or
05-generate-answer.py) over every golden question; each line needs
question, answer_text, cited_doc_ids and abstained.

## Logic flow

Golden set (aim for 50-100 items covering pilot corpus and standards)
→ run the app over the golden questions → score with the evaluation
script → fix the top failures (chunking, rerank, prompts) → re-run →
pilot group → 4-week pilot with weekly feedback reports → go/no-go.

## Configuration

- Prompt labels carry "(example — replace with your value)".
- Default golden set path: /opt/airag/eval/golden.json.
- Default eval report path: /opt/airag/eval/eval-report.md.
- Default pilot group: AIRAG-Pilot.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |
| v2 | Fix pass: pilot feedback report temp SQL file deleted in try/finally (including on failure); prompt labels carry "(example — replace with your value)"; other scripts: labels only. No TESTING invalidations. |

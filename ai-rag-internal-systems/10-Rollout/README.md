# 10 — Rollout

## Purpose

This phase has no new scripts of its own. Rollout reuses the existing
pipeline: runbooks, user guides, helpdesk training, quarterly reviews
and project handover are manual deliverables, and the one scriptable
rollout step (phase-2 corpus expansion) is already covered by the
phase-03 ingest script.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| (none in this folder) | P10-04 via 03-ingest-corpus.py | Re-run 03-ingest-corpus.py with the phase-2 source directory and corpus name `phase2`, then re-run 03-map-documents-to-clauses.py for the new docs. |

## Inputs

See 03-Internal-ingestion/README.md for the ingest script prompts
(source_directory, corpus_name, chunk_size, chunk_overlap,
output_jsonl_path, ingest_mode).

## Logic flow

1. Confirm the pilot go/no-go decision (P9-07).
2. Ingest phase-2 sources with 03-ingest-corpus.py.
3. Re-map the new docs to clauses and send the uncertain mappings for
   doc-owner review.
4. Publish runbooks and the user guide, train the helpdesk, schedule
   the quarterly model/corpus review, hand over to ops.

## Configuration

- Phase-2 corpus name: `phase2` (prompted at ingest time).
- No services, ports or credentials are introduced by this phase.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

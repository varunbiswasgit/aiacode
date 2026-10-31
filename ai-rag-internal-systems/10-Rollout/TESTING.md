# Testing — 10 Rollout

Manual test cases. Run each case, then tick the Pass column. This
phase has no scripts of its own; the case below exercises the reused
phase-03 ingest script for phase-2 expansion.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Phase-2 sample source directory, pilot corpus already ingested | Run 03-ingest-corpus.py with the phase-2 directory and corpus name phase2, then 03-map-documents-to-clauses.py on the new chunks | New chunks carry corpus phase2, mappings file includes the new docs, uncertain mappings marked needs_review | ☐ |

# Testing — 03 Internal ingestion

Manual test cases. Run each case, then tick the Pass column. Use a
small sample corpus (5-10 documents) for every case.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Sample SharePoint site or SMB share, read-only test account | Run 03-connect-document-sources.py for that source type (sample default /opt/airag/source-sample.json) | Script reports reachable, sample listing saved at the prompted path (default /opt/airag/source-sample.json) with mode 600, no token in the saved file; for SMB the mode-600 auth file is deleted afterwards and the credential never appears on the command line | ☐ |
| Sample corpus with PDF, DOCX, TXT, HTML files | Run 03-parse-documents.py with test input and output paths | Output JSONL has one record per chunk, each with doc_id, source_path, chunk_index, text | ☐ |
| Chunk JSONL from the parse test | Run 03-extract-metadata.py with a test doc-ID regex | Every output chunk carries a metadata object with title, owner, department and the filename-derived doc ID | ☐ |
| App test VM, venv present | Run 03-install-ocr.sh | tesseract --version and ocrmypdf --version both work | ☐ |
| App test VM, ingest script present | Run 03-schedule-ingestion.sh with test schedules, answering the extra_write_dirs prompt with every directory the ingest run writes to (default example /data/staging) | systemctl list-timers shows both airag timers enabled; the service uses ProtectSystem=strict with ReadWritePaths covering the project dir, log dir and the prompted extra_write_dirs, so a test output path not entered there is not writable by the scheduled run | ☐ |
| Sample corpus directory | Run 03-ingest-corpus.py in full mode, then again in incremental mode | First run writes all chunks, second run adds zero chunks for unchanged files | ☐ |
| Chunk JSONL from the ingest test | Run 03-validate-ingest.py with sample size 5, answer each prompt | Script prints a pass rate and lists any failed doc IDs | ☐ |
| Chunk JSONL plus a small clause JSONL | Run 03-map-documents-to-clauses.py | Mappings file contains confirmed rows for cited clauses and needs_review rows for keyword-only matches | ☐ |
| Chunk JSONL from the ingest test, test Postgres database | Run 03-build-chunks-fts.py in upsert mode, then again, then in rebuild mode with confirmation | chunks_fts holds one row per valid chunk with a GIN-indexed english tsvector; the second upsert adds no duplicates; rows missing doc_id, chunk_index or text are skipped and counted; rebuild recreates the table only after confirmation | ☐ |

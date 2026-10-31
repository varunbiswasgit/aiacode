# Testing — 03 Internal ingestion

Manual test cases. Run each case, then tick the Pass column. Use a
small sample corpus (5-10 documents) for every case.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Sample SharePoint site or SMB share, read-only test account | Run 03-connect-document-sources.py for that source type | Script reports reachable, sample listing saved, no token in the saved file | ☐ |
| Sample corpus with PDF, DOCX, TXT, HTML files | Run 03-parse-documents.py with test input and output paths | Output JSONL has one record per chunk, each with doc_id, source_path, chunk_index, text | ☐ |
| Chunk JSONL from the parse test | Run 03-extract-metadata.py with a test doc-ID regex | Every output chunk carries a metadata object with title, owner, department and the filename-derived doc ID | ☐ |
| App test VM, venv present | Run 03-install-ocr.sh | tesseract --version and ocrmypdf --version both work | ☐ |
| App test VM, ingest script present | Run 03-schedule-ingestion.sh with test schedules | systemctl list-timers shows both airag timers enabled | ☐ |
| Sample corpus directory | Run 03-ingest-corpus.py in full mode, then again in incremental mode | First run writes all chunks, second run adds zero chunks for unchanged files | ☐ |
| Chunk JSONL from the ingest test | Run 03-validate-ingest.py with sample size 5, answer each prompt | Script prints a pass rate and lists any failed doc IDs | ☐ |
| Chunk JSONL plus a small clause JSONL | Run 03-map-documents-to-clauses.py | Mappings file contains confirmed rows for cited clauses and needs_review rows for keyword-only matches | ☐ |

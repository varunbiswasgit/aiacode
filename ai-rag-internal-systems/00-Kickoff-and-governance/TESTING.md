# Testing — 00 Kickoff and governance

Manual test cases. Run each case, then tick the Pass column. Every
prompt answer below uses test values, never real company credentials.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Git server reachable, empty test repo path | Run 00-create-project-repo.sh, answer prompts with test host and path, confirm the bare-repo step, complete the clone | Script clones successfully, branches main, develop, docs exist, docs, prompts, sql, eval folders present | ☐ |

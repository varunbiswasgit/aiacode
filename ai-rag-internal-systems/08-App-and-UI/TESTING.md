# Testing — 08 App and UI

Manual test cases. Run each case, then tick the Pass column. Use test
connection values and a test database.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| App test node, Qdrant and model server reachable | Run 08-app-main.py, answer prompts with test values | App serves on the chosen port, home page loads, admin page shows Qdrant and model reachable | ☐ |
| App running, test audit database | Submit a review decision via the review form and a feedback entry via the feedback endpoint | review_decisions and feedback rows appear in the database, no document content changes | ☐ |
| App test node, app installed | Run 08-create-systemd-services.sh, then reboot the test node | airag-app service is active after reboot, app answers on its port | ☐ |

# Testing — 08 App and UI

Manual test cases. Run each case, then tick the Pass column. Use test
connection values and a test database.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| App test node, Qdrant and model server reachable, SSO config and user-groups map present | Start 08-app-main.py, answer the new prompts (incl. standards/docs database names, SSO config path, user-groups map path, session TTL, AIRAG_RETRIEVE_TOP_K, AIRAG_RRF_K) with test values, then log in via the login step (OIDC or LDAP per the SSO config) | After login, the app serves on the chosen port with a session cookie, home page loads, admin page shows Qdrant and model reachable; an unauthenticated request is redirected to /login (HTML) or answered 401 (POST JSON) | ☐ |
| App running with a logged-in session and an existing suggestion row in the test standards database | Submit a review decision via that suggestion's own review form (one form per suggestion) and a feedback entry via the feedback endpoint | review_decisions row appears with the suggestion's real document_id and clause_id, the session username as reviewer and an approved/rejected decision, feedback row appears, no document content changes | ☐ |
| App test node, app installed, SSO config and user-groups map present | Run 08-create-systemd-services.sh, answering the Step 1b env-file prompts, then reboot the test node | airag-app service is active after reboot with no controlling terminal, app answers on its port; the EnvironmentFile is mode 600 and the unit references it without a dash prefix (so a missing file fails the unit explicitly rather than silently). NOTE: the service needs the SSO config and user-groups map the app reads at startup. Step 1b now writes every AIRAG_* variable the app reads to the env file, so no manual env-file edit is needed; verify the env file contains all 31 variables (see README) | ☐ |

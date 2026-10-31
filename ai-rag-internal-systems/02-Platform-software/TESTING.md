# Testing — 02 Platform software

Manual test cases. Run each case, then tick the Pass column. Use test
VMs matching the target node roles.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Test VM, no Docker | Run 02-install-docker.sh, answer no to the registry mirror | docker run hello-world succeeds, docker compose version prints, service account in docker group | ☐ |
| Test VM, no Python 3.12 | Run 02-install-python.sh with a test venv path | venv python --version shows 3.12, pip works inside the venv | ☐ |
| Test VM, repo exists on test Git server | Run 02-clone-project-repo.sh with test server and path | Project directory contains a git checkout on the requested branch | ☐ |
| DB test VM, no Postgres | Run 02-install-postgres.sh | postgresql service active, SELECT version returns 16 | ☐ |
| DB test VM, Postgres installed | Run 02-enable-pgvector.sh for the test database | CREATE EXTENSION vector succeeds, extension listed by dx | ☐ |
| DB test VM, Postgres plus pgvector | Run 02-create-databases.sql via psql, answer prompts with test names | Three databases exist, app user can connect, app user has no superuser rights | ☐ |
| DB test VM, Docker installed | Run 02-install-qdrant.sh with test ports | Container running, curl to the health endpoint returns version JSON | ☐ |
| DB test VM, Docker installed | Run 02-install-redis.sh with a test port | redis-cli ping returns PONG, data persists after container restart | ☐ |
| GPU test VM, driver working | Run 02-install-ollama.sh with a test port | ollama --version works, API tags endpoint answers | ☐ |
| GPU test VM, venv present | Run 02-install-vllm.sh | python -c import vllm prints a version, nvidia-smi lists the GPU | ☐ |
| App test VM, app not yet running | Run 02-install-nginx-tls.sh, complete the CSR step with a test CA | nginx -t passes, HTTPS site proxies to the upstream port, HTTP redirects to HTTPS | ☐ |
| Test VM, Docker installed | Run 02-configure-log-aggregation.sh with test SIEM | Test logger message and container log arrive in the SIEM search | ☐ |
| All three test nodes up | Run 02-platform-smoke-test.sh from the app node with test hostnames | All checks print OK, script exits 0 | ☐ |

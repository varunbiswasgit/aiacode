# 00 — Kickoff and governance

## Purpose

Bootstrap the AI-RAG project repository so every later phase has one place
to store code, docs, prompts, SQL and evaluation data. This folder covers
the scriptable part of tracker phase 0; approvals, scope documents, RACI
charts and security reviews remain manual steps owned by the project
manager, IT security and QA.

## Scripts

| Script | Tracker tasks | What it does |
|---|---|---|
| 00-create-project-repo.sh | P0-12 | Creates the bare repo on the internal Git server, clones it locally, seeds main/develop/docs branches and the docs, prompts, sql, eval folders. |

## Inputs (prompts asked at runtime)

| Script | Prompts |
|---|---|
| 00-create-project-repo.sh | git_server_hostname, git_repository_path, local_clone_path, default_branch |

No secrets are collected by this script. No company values are hardcoded;
every hostname and path is prompted with a sensible default shown in
brackets.

## Logic flow

1. Prompt for the Git server, repository path, local clone path and
   default branch.
2. Print the server-side commands for creating the bare repository and
   wait for confirmation that it exists.
3. Clone (or pull, if already cloned) into the local path.
4. Seed the branch structure and standard folders, commit, and push.

## Configuration

- Default branch: `main` (prompted).
- Seeded folders: `docs`, `prompts`, `sql`, `eval`.
- Log file: `~/airag-create-repo.log`.

## Version history

| Version | Summary |
|---|---|
| v1 | Initial version |

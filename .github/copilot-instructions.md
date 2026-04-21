# Project Guidelines

Personal document repository + automation scripts. Copilot acts as a personal assistant over generated data files.

## Layout

- `scripts/` — Node/Playwright scripts that produce data files.
- `.github/skills/<name>/SKILL.md` — what each skill does, inputs, outputs.
- `.github/agents/*.agent.md` — custom agents (e.g. Assistant).
- `.github/prompts/*.prompt.md` — reusable slash-command prompts.
- `data/<source>/` — raw outputs from skills (gitignored). Treat as read-only.
- `notes/YYYY-MM-DD-<topic>.md` — assistant-generated summaries. Always link back to the source data file.
- `.auth/` — Playwright persistent browser profiles (gitignored).

## Code style

- Node ESM (`.mjs`). No TypeScript.
- Playwright scripts use `chromium.launchPersistentContext` so login survives between runs. Never hardcode credentials.
- Scripts are idempotent: each run writes a new timestamped file, never overwrites.

## Build and test

- Install: `npm install` (also runs `playwright install chromium`).
- Run a skill: `npm run scrape:outlook -- --days 3`.

## Conventions

- Never mutate remote state (no sending email, no deleting messages, no API writes).
- Before running a skill, read its `SKILL.md`.
- Digest notes group by **Act today / This week / FYI / Noise**.

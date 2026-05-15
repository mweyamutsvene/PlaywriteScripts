---
description: "Teams catch-up: scrape a supplied list of chats/channels and produce a prioritized digest under notes/."
argument-hint: "days back (default 1)"
agent: "Assistant"
---

Run my Teams catch-up routine for the last ${input:days:1} day(s).

Steps:

1. Read [.github/skills/teams-scrape/SKILL.md](../skills/teams-scrape/SKILL.md) so you know what the scraper produces.
2. Look for `teams-targets.json` at the repo root. If it's missing, ask me for the list of chats and channels to watch before continuing.
3. Run:
   ```pwsh
   npm run scrape:teams -- --days ${input:days:1}
   ```
   If a Chromium window opens and this is the first run, tell me to sign in to Teams; then wait.
4. Read the newest `data/teams/teams-*.json`.
5. Write a digest to `notes/YYYY-MM-DD-teams-digest.md` with:
   - Totals (targets, chats hit, messages, unique authors).
   - **Act today** — chat, who/what, one-line ask, suggested reply bullet.
   - **This week** — chat, one-line note.
   - **FYI** — one line each.
   - **Noise** — chats where nothing mattered (1 line each).
   - Footer linking the source JSON and MD files.
6. End your reply with the top 3 Act-today items and nothing else.

Do not invent messages. Do not send anything to Teams. Ask before running any other command.

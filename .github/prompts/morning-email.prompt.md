---
description: "Morning email catch-up: scrape Outlook for the last N days and produce a prioritized digest under notes/."
argument-hint: "days back (default 3)"
agent: "Assistant"
---

Run my email catch-up routine for the last ${input:days:3} days.

Steps:

1. Read [.github/skills/outlook-scrape/SKILL.md](../skills/outlook-scrape/SKILL.md) so you know what the scraper produces.
2. Run:
   ```pwsh
   npm run scrape:outlook -- --days ${input:days:3}
   ```
   If a Chromium window opens and this is the first run, tell me to sign in to Outlook; then wait.
3. Read the newest `data/outlook/inbox-*.json`.
4. Write a digest to `notes/YYYY-MM-DD-email-digest.md` with:
   - Totals (messages, unread, unique senders).
   - **Act today** — sender, subject, one-line ask, suggested reply bullet.
   - **This week** — sender, subject, one-line note.
   - **FYI** — one line each.
   - **Noise** — counts grouped by sender domain.
   - Footer linking the source JSON and MD files.
5. End your reply with the top 3 Act-today items and nothing else.

Do not invent content. Do not modify the mailbox. Ask before running any other command.

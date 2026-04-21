# Playwrite Scripts — Personal Assistant Workspace

A VS Code workspace that doubles as:

1. A **document repository** for personal notes, digests, and scraped data.
2. A set of **skills** (Node/Playwright scripts) that produce data files.
3. A **Copilot assistant** (`.github/chatmodes/assistant.chatmode.md`) that runs
   routine tasks and writes summaries on demand.

## Quick start

```pwsh
npm install
```

Then, in VS Code Copilot Chat:

1. Select the **Assistant** custom agent in the chat agent picker.
2. Type `/morning-email` (or paste a custom request).

On the very first run Chromium opens — sign in to Outlook, then leave the window.
All subsequent runs reuse the saved profile in `.auth/outlook/`.

## Layout

```
scripts/                    Node scripts
  scrape-outlook.mjs        Outlook Web inbox scraper
  lib/outlook-dom.mjs       DOM helpers injected into the page
data/<source>/              Raw outputs (gitignored)
notes/                      Assistant-generated summaries
.github/
  copilot-instructions.md   Workspace-wide rules for Copilot
  agents/assistant.agent.md Custom agent persona
  prompts/morning-email.prompt.md
  skills/outlook-scrape/SKILL.md  Skill: inputs, outputs, procedure
.auth/                      Persistent browser profiles (gitignored)
```

## Adding a new skill

1. Create `scripts/<your-skill>.mjs`.
2. Add an `npm run` entry in `package.json`.
3. Document it at `.github/skills/<your-skill>/SKILL.md` with YAML frontmatter (`name`, `description`).
4. Optionally add a prompt at `.github/prompts/<name>.prompt.md` that chains the skill + a summary step.

## Roadmap / alternatives

- **Chrome extension path**: a MV3 extension that scrapes the same DOM from the real
  browser session (no separate login) and drops files into a local folder via the
  File System Access API. Useful if the Playwright persistent-profile login ever
  breaks due to conditional access. Tracked as a future skill.
- **Calendar digest** — same pattern against `outlook.office.com/calendar`.
- **Teams unread digest** — same pattern against Teams web.

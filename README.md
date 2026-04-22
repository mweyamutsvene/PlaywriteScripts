# Playwrite Scripts — Personal Assistant Workspace

A VS Code workspace that doubles as:

1. A **document repository** for personal notes, digests, and scraped data.
2. A set of **skills** (Node/Playwright scripts) that produce data files.
3. A **Copilot assistant** (`.github/agents/assistant.agent.md`) that runs
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
  scrape-teams.mjs          Teams chats + channels scraper (enterprise)
  lib/outlook-dom.mjs       DOM helpers injected into the page
  lib/teams-dom.mjs
data/<source>/              Raw outputs (gitignored)
notes/                      Assistant-generated summaries (gitignored)
.github/
  copilot-instructions.md   Workspace-wide rules for Copilot
  agents/assistant.agent.md Custom agent persona
  prompts/morning-email.prompt.md
  prompts/teams-catchup.prompt.md
  skills/outlook-scrape/SKILL.md
  skills/teams-scrape/SKILL.md
.auth/                      Persistent browser profiles (gitignored)
```

## Running the scrapers directly

```pwsh
# Outlook inbox, last 3 days → data/outlook/inbox-<stamp>.{json,md}
npm run scrape:outlook -- --days 3

# Teams: inspect the DOM first to confirm selectors match your tenant
npm run scrape:teams -- --inspect --debug

# Teams: by sidebar name (chat, group chat, or a team/channel shown in the left rail)
npm run scrape:teams -- --names "Standup,iOS Dev" --days 1 --debug

# Teams: from a targets file (see teams-targets.example.json)
copy teams-targets.example.json teams-targets.json   # edit to taste (gitignored)
npm run scrape:teams -- --targets teams-targets.json --days 3
```

### Target formats

`--targets <file>` points at a JSON file shaped like:

```json
{
  "targets": [
    "Team Standup",
    { "team": "iOS Engineering Community", "channel": "General" },
    { "name": "Nick Baudin", "label": "Nick (1:1)" }
  ]
}
```

- **String** — matched against any chat, group chat, team, or channel title in
  the left sidebar. Favors chats over teams when multiple items share a name.
- **`{ team, channel }`** — team-scoped: finds the team in the sidebar, expands
  it if collapsed, then matches `channel` **only within that team's subtree**.
  Required for channels with common names like `General`, since almost every
  team has one.
- **`{ name, label? }`** — same as the string form but lets you set a friendlier
  label used in the Markdown output.

Navigation tries the **left-rail tree first** (fast, exact, no search UI). If
that misses, it falls back to Teams' unified search. With `--debug` you'll see
which path was used and, on a miss, the top sidebar candidates with scores.

### Teams script flags

| Flag | Default | Purpose |
| --- | --- | --- |
| `--targets <file>` | `teams-targets.json` | JSON array of chat/channel names |
| `--names "A,B"` | — | Comma-separated names (alternative to `--targets`) |
| `--url <url>` | `https://teams.microsoft.com/v2/` | Override Teams URL |
| `--days <n>` | `3` | Look back N days |
| `--since <iso>` | — | Harvest messages newer than this ISO date |
| `--max <n>` | `200` | Max messages per target |
| `--headful` | `true` | Show the browser window |
| `--inspect` | — | Print DOM diagnostics and exit |
| `--debug` | — | Verbose `[debug]` traces through search/pick/harvest |
| `--channel <name>` | bundled chromium | Use installed Chrome/Edge (e.g. `--channel chrome`) — survives Entra device checks better |

`TEAMS_DEBUG=1` also enables debug mode.

### If Teams keeps asking you to sign in every run

Enterprise tenants with conditional access often invalidate Playwright's bundled
Chromium session. Try:

```pwsh
# Use your real Chrome install (profile lives in .auth/teams, not your main Chrome profile)
npm run scrape:teams -- --channel chrome --inspect --debug
```

The script also:

- sets a realistic Chrome user-agent,
- disables the `--enable-automation` flag and the `navigator.webdriver` bit,
- closes the browser cleanly on Ctrl+C so cookies are flushed to disk.

If SSO still re-prompts every run, the tenant is enforcing session binding to
a managed/compliant device — no scraper workaround exists; run it on a device
that satisfies the Intune/Entra compliance policy.

## Setting this up on another machine

```pwsh
git clone https://github.com/mweyamutsvene/PlaywriteScripts.git
cd PlaywriteScripts
npm install                                  # also runs `playwright install chromium`

# First run opens Chromium — sign in to Outlook and/or Teams, then leave the window.
# Login survives between runs via the .auth/ persistent profiles (gitignored).
npm run scrape:outlook -- --days 1
npm run scrape:teams -- --inspect --debug
```

Persistent profiles are per-machine and never committed. The `--debug` output on
a fresh machine is the fastest way to confirm selectors match the tenant before
running against a full target list.

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

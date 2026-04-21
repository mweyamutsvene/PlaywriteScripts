---
name: teams-scrape
description: "Scrape enterprise Microsoft Teams (teams.microsoft.com/v2) chats and channels via Playwright and write a timestamped JSON + Markdown dump. Use when the user asks to catch up on Teams, summarize chats/channels, or produce a Teams digest from a list of targets."
argument-hint: "--targets teams-targets.json --days N   (or --names \"A,B\")"
---

# teams-scrape

Scrape a **user-supplied list** of Teams chats and channels in the enterprise web app using a persistent Chromium profile so login survives between runs.

Both chats and channels are opened via Teams' built-in **Go to** command (Ctrl+Alt+G, fallback Ctrl+E), which searches across both surfaces, so one code path handles both.

## When to Use

- User asks to "catch me up on Teams", "summarize my chats/channels", or produce a Teams digest.
- User wants a digest scoped to a specific set of chats/channels they name.
- User provides a `teams-targets.json` file or a comma-separated list on the CLI.

## Procedure

1. Ensure deps are installed: `npm install` (first run only).
2. Confirm we have a target list. If not, ask the user. Supported shapes:
   - File at repo root: `teams-targets.json`
     ```json
     { "targets": [
         "Team Standup",
         { "team": "Platform Eng", "channel": "General" },
         { "name": "Nick Baudin", "label": "Nick (1:1)" }
     ] }
     ```
   - CLI flag: `--names "Team Standup,Platform Eng General,Nick Baudin"`
3. Run:
   ```pwsh
   npm run scrape:teams -- --days 3
   npm run scrape:teams -- --targets teams-targets.json --days 1
   npm run scrape:teams -- --names "Standup,Platform Eng General" --days 1
   ```
4. Two files are written under `data/teams/`:
   - `teams-<ISO-timestamp>.json`
   - `teams-<ISO-timestamp>.md`
5. Read the newest JSON and produce the digest requested by the user.

## Inputs

| flag | default | meaning |
|---|---|---|
| `--targets <file>` | — | JSON array or `{targets:[...]}` of targets |
| `--names "A,B"` | — | shorthand: comma-separated names passed to Teams Go-to |
| `--url <teams-url>` | `https://teams.microsoft.com/v2/` | Teams web URL |
| `--days N` | 3 | include messages on/after midnight N days ago |
| `--since YYYY-MM-DD` | — | explicit cutoff, overrides `--days` |
| `--max N` | 200 | per-target cap on messages captured |
| `--headful` | true | show the browser window |
| `--inspect` | — | open Teams, dump selector diagnostics for the open pane, exit |

## Target shapes

The script normalizes each target to `{ name, label? }` where `name` is the string typed into the Go-to box:

- `"Team Standup"` → `{ name: "Team Standup" }`
- `{ "team": "Eng", "channel": "General" }` → `{ name: "Eng General", label: "Eng / General" }`
- `{ "name": "Nick Baudin", "label": "Nick (1:1)" }` → passed through

## Output schema

```json
{
  "meta": { "generatedAt": "...", "cutoff": "...", "url": "...", "targetCount": 3,
            "params": {"days":3,"since":null,"max":200} },
  "targets": [ { "name": "Team Standup" } ],
  "results": [
    {
      "target": { "name": "Team Standup" },
      "status": "ok",
      "header": "Team Standup",
      "messageCount": 12,
      "messages": [
        { "index": 3, "isSystemMessage": false, "sender": "Hanks, Tommy",
          "text": "...", "timeISO": "2026-04-20T15:22:00.000Z",
          "timeLabel": "Today 10:22 AM" }
      ]
    }
  ]
}
```

Possible `status` values: `ok`, `not-found`, `error`.

## How it works

- [scripts/scrape-teams.mjs](../../../scripts/scrape-teams.mjs) launches Chromium with a persistent profile at `.auth/teams/` and opens `teams.microsoft.com/v2/`.
- For each target it presses **Ctrl+Alt+G**, types the name, picks the first matching suggestion (`[role="option"]` / `[data-tid^="suggestion"]` / `[data-tid^="search-result"]`), and waits for either the header to change or message rows to render.
- [scripts/lib/teams-dom.mjs](../../../scripts/lib/teams-dom.mjs) holds the multi-selector banks for the enterprise app, ported from a verified Chrome extension:
  - `CHAT_CONTAINER_SELECTORS`: `.fui-Chat`, `[data-tid="message-pane-list-surface"]`, `.ts-message-list-container`, `[data-tid="chat-pane"]`, `[role="main"] [data-is-scrollable="true"]`, `[data-tid="messageListContainer"]`, `[role="log"]`, and class-prefix matches.
  - `MESSAGE_SELECTORS`: `[data-tid="chat-pane-item"]` (primary — wraps every row including system), plus `[data-testid="message-wrapper"]`, `[data-tid="chat-pane-message"]`, `[class*="fui-unstable-ChatItem"]`, etc.
  - `TIMESTAMP_SELECTORS`: `[class*="fui-ChatMessage__timestamp"]`, `[data-tid="messageTimeStamp"]`, `time[datetime]`, `[datetime]`, class-prefix fallbacks.
  - `SENDER_SELECTORS`: `[data-tid="message-author-name"]`, `[class*="fui-ChatMessage__author"]`, `.ui-chat__message__author`, class-prefix fallbacks, plus a "LastName, FirstName" regex fallback.
  - `MESSAGE_TEXT_SELECTORS`: `[class*="fui-ChatMessage__body"]`, `.ui-chat__message__content`, `.message-body-content`, `[data-tid="messageBodyContent"]`, class-prefix fallbacks.
- Date-divider rows (`[class*="fui-Divider"]`) are skipped; control/system rows (`[class*="ChatControlMessage"]`) are captured with `isSystemMessage: true` and inherit the last seen timestamp.
- Timestamps parse "Today HH:MM AM/PM", "Yesterday HH:MM", bare "HH:MM AM/PM", "M/D/YYYY HH:MM", and ISO strings.
- Scroll harvesting: walks the pane scroll container upward (`scrollPaneUpBy`) until either a message older than the cutoff is seen, `MAX_MESSAGES` is hit, or scroll stagnates for 3 rounds.
- Deduped by `timeISO|sender|text[:80]`.

## Known limits

- Enterprise web app only (`teams.microsoft.com/v2`). Consumer Teams is not supported.
- First-run login waits up to 5 minutes.
- Target resolution depends on Teams' Go-to picking the same first result a user would — ambiguous names may open the wrong thread. Prefer full names or `Team Channel` strings.
- Message text only; no attachments, images, or reactions.
- When Teams ships DOM changes, run `npm run scrape:teams -- --inspect` and update the selector banks in `scripts/lib/teams-dom.mjs`.
---
name: teams-scrape
description: "Scrape selected Microsoft Teams chats (and, on enterprise Teams, channels) via Playwright and write a timestamped JSON + Markdown dump. Use when the user asks to catch up on Teams, summarize chats, or produce a Teams digest from a list of targets."
argument-hint: "--targets teams-targets.json --days N   (or --chats \"A,B\" --channels \"Team/Channel\")"
---

# teams-scrape

Scrape a **user-supplied list** of Teams chats and channels using a persistent Chromium profile so login survives between runs. Unlike the Outlook skill this one is list-driven: the user tells us which conversations matter.

## When to Use

- User asks to "catch me up on Teams", "summarize my chats", or "what did I miss in Teams".
- User wants a digest scoped to a specific set of chats/channels they name.
- User provides a `teams-targets.json` file or a comma-separated list on the CLI.

## Procedure

1. Ensure deps are installed: `npm install` (first run only).
2. Confirm we have a target list. If not, ask the user. Supported shapes:
   - File at repo root: `teams-targets.json`
     ```json
     { "targets": [
         { "type": "chat", "match": "Standup" },
         { "type": "chat", "match": "Nick Baudin" },
         { "type": "channel", "team": "Engineering", "channel": "General" }
     ] }
     ```
   - CLI flags: `--chats "Standup,Nick Baudin" --channels "Engineering/General"`
3. Run the scraper:
   ```pwsh
   npm run scrape:teams -- --days 3
   npm run scrape:teams -- --targets teams-targets.json --days 1
   npm run scrape:teams -- --chats "Standup,Nick" --days 1
   npm run scrape:teams -- --url https://teams.microsoft.com/v2/ --targets teams-targets.json
   ```
   - Default URL is `https://teams.live.com/v2/` (personal/consumer Teams).
   - For enterprise, pass `--url https://teams.microsoft.com/v2/`.
4. Two files are written under `data/teams/`:
   - `teams-<ISO-timestamp>.json`
   - `teams-<ISO-timestamp>.md`
5. Read the newest JSON and produce the digest requested by the user.

## Inputs

| flag | default | meaning |
|---|---|---|
| `--targets <file>` | — | JSON array or `{targets:[...]}` of chat/channel targets |
| `--chats "A,B"` | — | shorthand for chat targets (substring match on display name) |
| `--channels "Team/Channel,..."` | — | shorthand for channel targets (enterprise only) |
| `--url <teams-url>` | `https://teams.live.com/v2/` | Teams web URL |
| `--days N` | 3 | include messages on/after midnight N days ago |
| `--since YYYY-MM-DD` | — | explicit cutoff, overrides `--days` |
| `--max N` | 200 | per-target cap on messages captured |
| `--headful` | true | show the browser window |
| `--inspect` | — | print left-rail + channel selectors and exit (for debugging) |

## Output schema

```json
{
  "meta": { "generatedAt": "...", "cutoff": "...", "url": "...", "targetCount": 3, "params": {"days":3,"since":null,"max":200} },
  "targets": [ { "type": "chat", "match": "Standup" } ],
  "results": [
    {
      "target": { "type": "chat", "match": "Standup" },
      "status": "ok",
      "kind": "chat",
      "threadId": "19:...@thread.v2",
      "displayName": "Team Standup",
      "isGroup": true,
      "isMeeting": false,
      "header": "Team Standup",
      "messageCount": 12,
      "messages": [
        { "mid": "1776749315334", "author": "Nick", "bodyText": "...", "timeISO": "2026-04-20T...", "timeLabel": "Today at 11:28 PM." }
      ]
    }
  ]
}
```

Possible `status` values: `ok`, `not-found`, `open-failed`, `channels-not-available`, `channel-unimplemented`, `unknown-type`, `error`.

## How it works

- [scripts/scrape-teams.mjs](../../../scripts/scrape-teams.mjs) launches Chromium with a persistent profile at `.auth/teams/` and opens the Teams web URL.
- DOM helpers in [scripts/lib/teams-dom.mjs](../../../scripts/lib/teams-dom.mjs) target:
  - `[role="region"][data-tid="chat-list-layout"]` — left-rail chats region.
  - `[role="treeitem"][aria-level="2"]` containing `[data-tid="chat-list-item"]` — one per chat; id `chat-list-item_<threadId>`.
  - `.virtual-tree-list-scroll-container` — virtualized chat-list scroller.
  - `[data-tid="chat-pane-message"]` with `data-mid=<epoch ms>` — one per rendered message in the right pane.
  - `#content-<mid>`, `#author-<mid>`, `#timestamp-<mid>` — body/author/time for each message (`timestamp` is a `<time datetime="...">` element).
  - `[data-tid="chat-title"]` — header of the open chat.
- Flow: walk left rail to index every chat → match each target by substring on `displayName` → click the chat row → scroll the message pane **up** (older) until cutoff is reached → collect messages → move to next target.
- Channels: only inspected on enterprise Teams. If no channel selectors are found the target is marked `channels-not-available`. Full channel navigation is not yet implemented — the skill notes this explicitly so the digest can fall back to chats only.

## Known limits

- Consumer `teams.live.com` has no channels.
- Group-chat authorship: uses `data-tid="message-author-name"` when present; some system messages have no author.
- Message body is text only (no attachments/images).
- Date filtering uses `data-mid` (ms epoch) when available; otherwise `<time datetime>`. Older messages without either fall back to being included.
- First-run login window stays open up to 5 minutes.

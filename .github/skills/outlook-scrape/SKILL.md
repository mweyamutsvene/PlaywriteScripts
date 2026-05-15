---
name: outlook-scrape
description: "Scrape the Outlook Web inbox via Playwright and write a timestamped JSON + Markdown dump. Use when the user asks to fetch email, catch up on inbox, build an email digest, or scrape Outlook."
argument-hint: "--days N  (or --since YYYY-MM-DD, --max N, --headful)"
---

# outlook-scrape

Scrape the Outlook Web inbox using a persistent Chromium profile so login survives between runs.

## When to Use

- User asks to fetch/pull/scrape/download email.
- User asks for a morning/weekend email catch-up or digest.
- User wants raw inbox data before an LLM summary step.

## Procedure

1. Ensure deps are installed: `npm install` (first run only).
2. Run the scraper (headful so the user can sign in on first run):
   ```pwsh
   npm run scrape:outlook -- --days 3
   ```
   Variants:
   ```pwsh
   node scripts/scrape-outlook.mjs --since 2026-04-18 --max 300
   node scripts/scrape-outlook.mjs --days 7
   ```
3. Two files are written under `data/outlook/`:
   - `inbox-<ISO-timestamp>.json` — canonical record (see schema below).
   - `inbox-<ISO-timestamp>.md` — same content, human-readable.
4. Read the newest JSON and produce the digest requested by the user.

## Inputs

| flag | default | meaning |
|---|---|---|
| `--days N` | 3 | include messages on/after midnight N days ago |
| `--since YYYY-MM-DD` | — | explicit cutoff, overrides `--days` |
| `--max N` | 500 | hard cap on messages captured |
| `--headful` | true | show the browser window |

## Output schema

```json
{
  "meta": { "generatedAt": "...", "cutoff": "...", "count": 42, "params": {"days":3,"since":null,"max":500} },
  "messages": [
    {
      "convId": "AAQk...", "id": "AQAA...", "isUnread": true,
      "subject": "...", "fromName": "...", "fromEmail": "...",
      "to": ["..."], "cc": [],
      "dateTitle": "Mon 4/20/2026 5:16 PM",
      "dateISO": "2026-04-20T17:16:00.000Z",
      "bodyText": "..."
    }
  ]
}
```

## How it works

- [scripts/scrape-outlook.mjs](../../../scripts/scrape-outlook.mjs) launches Chromium with a persistent profile at `.auth/outlook/`.
- DOM helpers in [scripts/lib/outlook-dom.mjs](../../../scripts/lib/outlook-dom.mjs) target:
  - `[aria-label="Message list"] .customScrollBar[data-is-scrollable="true"]` — the virtualized scroller.
  - `[role="option"][data-convid]` — each conversation row.
  - `[aria-label="Reading Pane"] [aria-label="Message body"]` — the opened message body.
- The script scrolls to harvest IDs, dedupes by `data-convid`, stops at the cutoff, scrolls back to top, then opens each conversation to capture the body.

## Known limits

- Scrapes the **currently selected folder** (default Inbox). Switch folders manually before running if you need Archive/Sent.
- Outlook groups by conversation; each record is the newest message in that conversation.
- Text-only body capture. No attachments, no inline images.
- First-run login window stays open up to 5 minutes.

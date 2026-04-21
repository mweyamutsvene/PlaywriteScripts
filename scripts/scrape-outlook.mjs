#!/usr/bin/env node
// Scrape Outlook Web inbox via Playwright, using a persistent Chromium profile
// so you only have to sign in once. Writes one JSON + one Markdown file per run.
//
// Usage:
//   node scripts/scrape-outlook.mjs --days 3
//   node scripts/scrape-outlook.mjs --since 2026-04-18
//   node scripts/scrape-outlook.mjs --max 200 --headful
//
// On first run a Chromium window opens; sign in to Outlook, then leave it.
// The script waits for the message list to appear before scraping.

import { chromium } from 'playwright';
import { mkdir, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  collectVisibleItems,
  scrollBy,
  scrollToTop,
  readOpenMessage,
} from './lib/outlook-dom.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const AUTH_DIR = path.join(ROOT, '.auth', 'outlook');
const DATA_DIR = path.join(ROOT, 'data', 'outlook');

const OUTLOOK_URL = 'https://outlook.office.com/mail/';

// ---------- args ----------
const args = parseArgs(process.argv.slice(2));
const DAYS = args.days != null ? Number(args.days) : (args.since ? null : 3);
const SINCE = args.since ? new Date(args.since) : null;
const MAX_ITEMS = args.max ? Number(args.max) : 500;
const HEADFUL = args.headful ?? true; // default headful so user can sign in
const SCROLL_STEP = 800;
const SCROLL_SETTLE_MS = 350;
const OPEN_SETTLE_MS = 700;

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (!a.startsWith('--')) continue;
    const key = a.slice(2);
    const next = argv[i + 1];
    if (next && !next.startsWith('--')) { out[key] = next; i++; }
    else out[key] = true;
  }
  return out;
}

// ---------- date parsing ----------
// Outlook shows relative titles like "Mon 4/20/2026 5:16 PM" in row metadata.
function parseOutlookDate(s) {
  if (!s) return null;
  const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) return null;
  let [, mo, d, y, hh, mm, ap] = m;
  hh = Number(hh); mm = Number(mm);
  if (/pm/i.test(ap) && hh !== 12) hh += 12;
  if (/am/i.test(ap) && hh === 12) hh = 0;
  return new Date(Number(y), Number(mo) - 1, Number(d), hh, mm);
}

function cutoffDate() {
  if (SINCE) return SINCE;
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - DAYS);
  return d;
}

// ---------- main ----------
(async () => {
  await mkdir(AUTH_DIR, { recursive: true });
  await mkdir(DATA_DIR, { recursive: true });

  const isFirstRun = !existsSync(path.join(AUTH_DIR, 'Default'));
  const context = await chromium.launchPersistentContext(AUTH_DIR, {
    headless: !HEADFUL,
    viewport: { width: 1400, height: 1000 },
  });
  const page = context.pages()[0] ?? await context.newPage();

  console.log(`→ Navigating to ${OUTLOOK_URL}`);
  await page.goto(OUTLOOK_URL, { waitUntil: 'domcontentloaded' });

  if (isFirstRun) {
    console.log('→ First run: sign in to Outlook in the opened window. Waiting up to 5 minutes...');
  }

  // Wait for the inbox list to render (survives login flow).
  await page.waitForSelector('[aria-label="Message list"] [role="option"][data-convid]', { timeout: 5 * 60_000 });
  console.log('✓ Inbox detected.');

  const cutoff = cutoffDate();
  console.log(`→ Collecting messages on/after ${cutoff.toISOString()} (max ${MAX_ITEMS})`);

  // Scroll to top so we start from newest.
  await page.evaluate(scrollToTop);
  await page.waitForTimeout(SCROLL_SETTLE_MS);

  /** @type {Map<string, any>} */
  const seen = new Map();
  let reachedCutoff = false;
  let stagnantScrolls = 0;
  let lastScrollTop = -1;

  while (!reachedCutoff && seen.size < MAX_ITEMS && stagnantScrolls < 3) {
    const snap = await page.evaluate(collectVisibleItems);
    for (const it of snap.items) {
      if (seen.has(it.convId)) continue;
      const parsed = parseOutlookDate(it.dateTitle);
      const rec = { ...it, dateISO: parsed ? parsed.toISOString() : null };
      seen.set(it.convId, rec);
      if (parsed && parsed < cutoff) {
        reachedCutoff = true;
      }
    }

    if (reachedCutoff || seen.size >= MAX_ITEMS) break;

    const scrollRes = await page.evaluate(scrollBy, SCROLL_STEP);
    if (!scrollRes?.ok) {
      console.warn('! Scroller not found; stopping.');
      break;
    }
    if (scrollRes.after === lastScrollTop || scrollRes.atBottom) {
      stagnantScrolls++;
    } else {
      stagnantScrolls = 0;
      lastScrollTop = scrollRes.after;
    }
    await page.waitForTimeout(SCROLL_SETTLE_MS);
  }

  // Keep only items within cutoff & sort newest-first.
  let items = [...seen.values()]
    .filter(it => it.dateISO && new Date(it.dateISO) >= cutoff)
    .sort((a, b) => new Date(b.dateISO) - new Date(a.dateISO))
    .slice(0, MAX_ITEMS);

  console.log(`→ ${items.length} messages within window. Opening each to capture body...`);

  // Reset scroller to the top so virtualized rows for the newest items remount
  // before we try to click them.
  await page.evaluate(scrollToTop);
  await page.waitForTimeout(SCROLL_SETTLE_MS);

  // Open each message one at a time to read body. Outlook keeps convId in DOM id.
  const results = [];
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    try {
      // Row may have been virtualized away; scroll until it exists.
      const opened = await openConversation(page, it);
      if (!opened) {
        console.warn(`  [${i + 1}/${items.length}] Could not find row for ${it.convId}; skipping.`);
        continue;
      }
      await page.waitForTimeout(OPEN_SETTLE_MS);
      // Wait for body to be present for this conversation.
      await page.waitForFunction(() => {
        const rp = document.querySelector('[aria-label="Reading Pane"]');
        return !!rp?.querySelector('[aria-label="Message body"]');
      }, { timeout: 15_000 }).catch(() => {});
      const msg = await page.evaluate(readOpenMessage);
      const parsedDate = parseOutlookDate(msg?.dateTitle) || parseOutlookDate(it.dateTitle);
      results.push({
        convId: it.convId,
        id: it.id,
        isUnread: it.isUnread,
        subject: msg?.subject ?? deriveSubjectFromAria(it.ariaLabel),
        fromName: msg?.fromName ?? it.senderName,
        fromEmail: msg?.fromEmail ?? it.senderEmail,
        to: msg?.to ?? [],
        cc: msg?.cc ?? [],
        dateTitle: msg?.dateTitle ?? it.dateTitle,
        dateISO: parsedDate ? parsedDate.toISOString() : null,
        bodyText: msg?.bodyText ?? null,
      });
      if ((i + 1) % 10 === 0) console.log(`  ${i + 1}/${items.length} captured`);
    } catch (err) {
      console.warn(`  error on ${it.convId}:`, err.message);
    }
  }

  // ---------- write output ----------
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const jsonPath = path.join(DATA_DIR, `inbox-${stamp}.json`);
  const mdPath = path.join(DATA_DIR, `inbox-${stamp}.md`);
  const meta = {
    generatedAt: new Date().toISOString(),
    cutoff: cutoff.toISOString(),
    count: results.length,
    params: { days: DAYS, since: SINCE?.toISOString() ?? null, max: MAX_ITEMS },
  };
  await writeFile(jsonPath, JSON.stringify({ meta, messages: results }, null, 2), 'utf8');
  await writeFile(mdPath, toMarkdown(meta, results), 'utf8');

  console.log(`✓ Wrote ${jsonPath}`);
  console.log(`✓ Wrote ${mdPath}`);

  await context.close();
})().catch(err => {
  console.error(err);
  process.exit(1);
});

// ---------- helpers ----------

async function openConversation(page, it) {
  const selector = `[aria-label="Message list"] [role="option"][data-convid="${cssEscape(it.convId)}"]`;

  // Try up to 10 scroll passes to bring the row into view, then click.
  for (let attempt = 0; attempt < 10; attempt++) {
    const handle = await page.$(selector);
    if (handle) {
      await handle.scrollIntoViewIfNeeded().catch(() => {});
      await handle.click({ timeout: 5000 }).catch(() => {});
      return true;
    }
    // Row isn't mounted. Estimate where it should be based on posInSet and
    // scroll directly there; fall back to small steps.
    const moved = await page.evaluate((pos) => {
      const scroller =
        document.querySelector('[aria-label="Message list"] .customScrollBar[data-is-scrollable="true"]') ||
        [...document.querySelectorAll('[aria-label="Message list"] .customScrollBar')].find(s => s.querySelector('[role="option"]'));
      if (!scroller) return { ok: false };
      const anyOpt = document.querySelector('[aria-label="Message list"] [role="option"][aria-posinset]');
      const total = anyOpt ? Number(anyOpt.getAttribute('aria-setsize')) : null;
      if (pos && total && total > 0) {
        const target = Math.max(0, Math.floor((pos - 2) / total * scroller.scrollHeight));
        scroller.scrollTop = target;
      } else {
        scroller.scrollTop = scroller.scrollTop + 800;
      }
      return { ok: true, scrollTop: scroller.scrollTop, scrollHeight: scroller.scrollHeight };
    }, it.posInSet);
    if (!moved?.ok) break;
    await page.waitForTimeout(250);
  }
  return false;
}

function cssEscape(s) {
  return String(s).replace(/["\\]/g, '\\$&');
}

function deriveSubjectFromAria(label) {
  // aria-label shape: "Unread <sender> <subject> <day> <time> <preview>..."
  // Best-effort only; reading-pane subject is preferred.
  return label?.split(/\s{2,}/)[1] ?? null;
}

function toMarkdown(meta, messages) {
  const lines = [];
  lines.push(`# Outlook inbox dump`);
  lines.push('');
  lines.push(`- Generated: ${meta.generatedAt}`);
  lines.push(`- Cutoff: ${meta.cutoff}`);
  lines.push(`- Messages: ${meta.count}`);
  lines.push('');
  for (const m of messages) {
    lines.push(`## ${m.subject ?? '(no subject)'}`);
    lines.push('');
    lines.push(`- **From:** ${m.fromName ?? ''}${m.fromEmail ? ` <${m.fromEmail}>` : ''}`);
    if (m.to?.length) lines.push(`- **To:** ${m.to.join(', ')}`);
    if (m.cc?.length) lines.push(`- **Cc:** ${m.cc.join(', ')}`);
    lines.push(`- **Date:** ${m.dateTitle ?? m.dateISO ?? ''}`);
    lines.push(`- **Unread:** ${m.isUnread ? 'yes' : 'no'}`);
    lines.push(`- **ConvId:** \`${m.convId}\``);
    lines.push('');
    lines.push('```text');
    lines.push((m.bodyText ?? '(no body captured)').trim());
    lines.push('```');
    lines.push('');
  }
  return lines.join('\n');
}

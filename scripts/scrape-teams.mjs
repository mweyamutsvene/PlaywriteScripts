#!/usr/bin/env node
// Scrape enterprise Microsoft Teams (teams.microsoft.com/v2) for a supplied
// list of chats and channels. Uses Playwright with a persistent Chromium
// profile so login survives runs.
//
// Targets are resolved via the built-in "Go to" command (Ctrl+Alt+G) which
// searches chats AND channels with a single query, so one code path handles
// both surfaces.
//
// Usage:
//   node scripts/scrape-teams.mjs --targets teams-targets.json --days 3
//   node scripts/scrape-teams.mjs --names "Standup,Platform Eng General" --days 1
//   node scripts/scrape-teams.mjs --inspect

import { chromium } from 'playwright';
import { mkdir, writeFile, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  readHeader,
  collectPaneMessages,
  scrollPaneUpBy,
  inspectPane,
} from './lib/teams-dom.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const AUTH_DIR = path.join(ROOT, '.auth', 'teams');
const DATA_DIR = path.join(ROOT, 'data', 'teams');

const args = parseArgs(process.argv.slice(2));
const TEAMS_URL = args.url ?? 'https://teams.microsoft.com/v2/';
const DAYS = args.days != null ? Number(args.days) : (args.since ? null : 3);
const SINCE = args.since ? new Date(args.since) : null;
const MAX_MESSAGES = args.max ? Number(args.max) : 200;
const HEADFUL = args.headful ?? true;
const INSPECT = !!args.inspect;
const PANE_SETTLE_MS = 900;
const MAX_SCROLL_ROUNDS = 40;

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

function cutoffDate() {
  if (SINCE) return SINCE;
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - DAYS);
  return d;
}

async function loadTargets() {
  if (args.targets) {
    const raw = await readFile(path.resolve(process.cwd(), args.targets), 'utf8');
    const parsed = JSON.parse(raw);
    const list = Array.isArray(parsed) ? parsed : parsed.targets;
    if (!Array.isArray(list)) throw new Error('--targets file must be an array or { targets: [...] }');
    return list.map(normalizeTarget);
  }
  if (typeof args.names === 'string') {
    return args.names.split(',').map(s => s.trim()).filter(Boolean).map(n => ({ name: n }));
  }
  const def = path.join(ROOT, 'teams-targets.json');
  if (existsSync(def)) {
    const raw = await readFile(def, 'utf8');
    const parsed = JSON.parse(raw);
    const list = Array.isArray(parsed) ? parsed : parsed.targets ?? [];
    return list.map(normalizeTarget);
  }
  return [];
}

function normalizeTarget(t) {
  if (typeof t === 'string') return { name: t };
  if (t && typeof t === 'object') {
    if (t.name) return { name: String(t.name), label: t.label };
    if (t.team && t.channel) return { name: `${t.team} ${t.channel}`, label: `${t.team} / ${t.channel}` };
    if (t.match) return { name: String(t.match) };
  }
  throw new Error(`Invalid target: ${JSON.stringify(t)}`);
}

(async () => {
  await mkdir(AUTH_DIR, { recursive: true });
  await mkdir(DATA_DIR, { recursive: true });

  const isFirstRun = !existsSync(path.join(AUTH_DIR, 'Default'));
  const context = await chromium.launchPersistentContext(AUTH_DIR, {
    headless: !HEADFUL,
    viewport: { width: 1400, height: 1000 },
  });
  const page = context.pages()[0] ?? await context.newPage();

  console.log(`â†’ Navigating to ${TEAMS_URL}`);
  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });

  if (isFirstRun) {
    console.log('â†’ First run: sign in to Teams in the opened window. Waiting up to 5 minutes...');
  }

  await page.waitForSelector(
    '[data-tid="app-bar-wrapper"], [role="main"], [data-tid="chat-pane"]',
    { timeout: 5 * 60_000 }
  );
  await page.waitForTimeout(1500);
  console.log('âœ“ Teams shell detected.');

  if (INSPECT) {
    const info = await page.evaluate(inspectPane);
    console.log(JSON.stringify(info, null, 2));
    await context.close();
    return;
  }

  const targets = await loadTargets();
  if (!targets.length) {
    console.error('! No targets. Provide --targets <file>, --names "A,B", or create teams-targets.json at repo root.');
    console.error('  Format: { "targets": ["Team Standup", {"team":"Eng","channel":"General"}] }');
    await context.close();
    process.exit(2);
  }
  console.log(`â†’ ${targets.length} target(s):`, targets.map(t => t.label || t.name).join(' | '));

  const cutoff = cutoffDate();
  console.log(`â†’ Cutoff: ${cutoff.toISOString()} (max ${MAX_MESSAGES}/target)`);

  const results = [];
  for (const t of targets) {
    console.log(`  â†’ ${t.label || t.name}`);
    try {
      const opened = await gotoTarget(page, t.name);
      if (!opened) {
        results.push({ target: t, status: 'not-found' });
        console.warn(`    âœ— could not open via Go-to`);
        continue;
      }
      await page.waitForFunction(
        () => document.querySelectorAll('[data-tid="chat-pane-item"], [data-tid="chat-pane-message"]').length > 0,
        { timeout: 20_000 }
      ).catch(() => {});
      await page.waitForTimeout(PANE_SETTLE_MS);

      const thread = await harvestOpenPane(page, cutoff);
      results.push({ target: t, status: 'ok', ...thread });
      console.log(`    âœ“ ${thread.messageCount} msgs (header: ${thread.header ?? '?'})`);
    } catch (err) {
      console.warn(`    error:`, err.message);
      results.push({ target: t, status: 'error', error: err.message });
    }
  }

  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const jsonPath = path.join(DATA_DIR, `teams-${stamp}.json`);
  const mdPath = path.join(DATA_DIR, `teams-${stamp}.md`);
  const meta = {
    generatedAt: new Date().toISOString(),
    cutoff: cutoff.toISOString(),
    url: TEAMS_URL,
    targetCount: targets.length,
    params: { days: DAYS, since: SINCE?.toISOString() ?? null, max: MAX_MESSAGES },
  };
  await writeFile(jsonPath, JSON.stringify({ meta, targets, results }, null, 2), 'utf8');
  await writeFile(mdPath, toMarkdown(meta, results), 'utf8');
  console.log(`âœ“ Wrote ${jsonPath}`);
  console.log(`âœ“ Wrote ${mdPath}`);

  await context.close();
})().catch(err => {
  console.error(err);
  process.exit(1);
});

async function gotoTarget(page, name) {
  const beforeHeader = await page.evaluate(readHeader).catch(() => ({ header: null }));

  await page.keyboard.press('Control+Alt+G').catch(() => {});
  await page.waitForTimeout(400);

  let opened = await pickFromSuggestions(page, name);
  if (!opened) {
    await page.keyboard.press('Escape').catch(() => {});
    await page.waitForTimeout(200);
    await page.keyboard.press('Control+E').catch(() => {});
    await page.waitForTimeout(500);
    opened = await pickFromSuggestions(page, name);
  }
  if (!opened) return false;

  return await page.waitForFunction(
    (prev) => {
      const hEl = document.querySelector(
        '[data-tid="chat-header-title"], [data-tid="chat-title"], [data-tid="channel-header-title"], [role="main"] h1, [role="main"] h2'
      );
      const h = hEl?.textContent?.trim() || null;
      const hasMsgs = document.querySelectorAll('[data-tid="chat-pane-item"], [data-tid="chat-pane-message"]').length > 0;
      return (h && h !== prev) || hasMsgs;
    },
    beforeHeader.header,
    { timeout: 15_000 }
  ).then(() => true).catch(() => false);
}

async function pickFromSuggestions(page, name) {
  const boxes = await page.$$(
    'input[type="text"], input[type="search"], [role="combobox"], [role="searchbox"], [contenteditable="true"]'
  );
  let input = null;
  for (const b of boxes) {
    if (await b.evaluate(el => el === document.activeElement).catch(() => false)) { input = b; break; }
  }
  if (!input) input = boxes[0];
  if (!input) return false;

  await input.click({ delay: 30 }).catch(() => {});
  await page.keyboard.press('Control+A').catch(() => {});
  await page.keyboard.press('Delete').catch(() => {});
  await page.keyboard.type(name, { delay: 20 });
  await page.waitForTimeout(900);

  const clicked = await page.evaluate((needle) => {
    const lower = String(needle).toLowerCase();
    const candidates = document.querySelectorAll(
      '[role="option"], [role="listitem"], [data-tid^="suggestion"], [data-tid^="search-result"]'
    );
    for (const el of candidates) {
      const t = (el.textContent || '').toLowerCase();
      if (t.includes(lower) || lower.split(/\s+/).every(tok => t.includes(tok))) {
        el.click();
        return true;
      }
    }
    const first = document.querySelector('[role="option"]');
    if (first) { first.click(); return true; }
    return false;
  }, name);

  if (!clicked) {
    await page.keyboard.press('Enter').catch(() => {});
  }
  await page.waitForTimeout(600);
  return true;
}

async function harvestOpenPane(page, cutoff) {
  const seen = new Map();
  let reachedOlder = false;
  let stagnant = 0;
  let lastTop = -1;

  for (let round = 0; round < MAX_SCROLL_ROUNDS && !reachedOlder && seen.size < MAX_MESSAGES && stagnant < 3; round++) {
    const snap = await page.evaluate(collectPaneMessages);
    for (const m of snap.msgs) {
      const key = `${m.timeISO || ''}|${m.sender || ''}|${(m.text || '').slice(0, 80)}`;
      if (seen.has(key)) continue;
      seen.set(key, m);
      if (m.timeISO && new Date(m.timeISO) < cutoff) reachedOlder = true;
    }
    if (reachedOlder || seen.size >= MAX_MESSAGES) break;

    const mv = await page.evaluate(scrollPaneUpBy, 1400);
    if (!mv?.ok) break;
    if (mv.atTop || mv.after === lastTop) stagnant++;
    else { stagnant = 0; lastTop = mv.after; }
    await page.waitForTimeout(PANE_SETTLE_MS);
  }

  const header = (await page.evaluate(readHeader)).header;
  const messages = [...seen.values()]
    .filter(m => !m.timeISO || new Date(m.timeISO) >= cutoff)
    .sort((a, b) => new Date(a.timeISO || 0) - new Date(b.timeISO || 0));
  return { header, messageCount: messages.length, messages };
}

function toMarkdown(meta, results) {
  const lines = [];
  lines.push('# Teams dump');
  lines.push('');
  lines.push(`- Generated: ${meta.generatedAt}`);
  lines.push(`- Cutoff: ${meta.cutoff}`);
  lines.push(`- URL: ${meta.url}`);
  lines.push(`- Targets: ${meta.targetCount}`);
  lines.push('');
  for (const r of results) {
    const label = r.target.label || r.target.name;
    lines.push(`## ${label}`);
    lines.push('');
    lines.push(`- Status: \`${r.status}\``);
    if (r.header) lines.push(`- Opened as: ${r.header}`);
    if (typeof r.messageCount === 'number') lines.push(`- Messages: ${r.messageCount}`);
    lines.push('');
    if (r.messages?.length) {
      for (const m of r.messages) {
        const tag = m.isSystemMessage ? ' _(system)_' : '';
        lines.push(`### ${m.sender ?? '(unknown)'} â€” ${m.timeISO ?? m.timeLabel ?? ''}${tag}`);
        lines.push('');
        lines.push('```text');
        lines.push((m.text ?? '').trim() || '(no text)');
        lines.push('```');
        lines.push('');
      }
    }
  }
  return lines.join('\n');
}

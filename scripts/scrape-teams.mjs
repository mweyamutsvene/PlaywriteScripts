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
  expandChannelReplies,
  clickSidebarTarget,
  ensureChatTabActive,
  scrollSidebarTree,
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
const DEBUG = !!args.debug || !!process.env.TEAMS_DEBUG;
const PANE_SETTLE_MS = 900;
const MAX_SCROLL_ROUNDS = 40;

function dbg(...a) { if (DEBUG) console.log('[debug]', ...a); }

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
    if (t.team && t.channel) return {
      name: `${t.team} ${t.channel}`,
      label: `${t.team} / ${t.channel}`,
      team: String(t.team),
      channel: String(t.channel),
    };
    if (t.match) return { name: String(t.match) };
  }
  throw new Error(`Invalid target: ${JSON.stringify(t)}`);
}

(async () => {
  await mkdir(AUTH_DIR, { recursive: true });
  await mkdir(DATA_DIR, { recursive: true });

  console.log('=== scrape-teams ===');
  console.log(`  node        ${process.version}`);
  console.log(`  cwd         ${process.cwd()}`);
  console.log(`  root        ${ROOT}`);
  console.log(`  auth dir    ${AUTH_DIR}`);
  console.log(`  data dir    ${DATA_DIR}`);
  console.log(`  url         ${TEAMS_URL}`);
  console.log(`  days/since  ${SINCE ? SINCE.toISOString() : DAYS + ' day(s)'}`);
  console.log(`  max msgs    ${MAX_MESSAGES}`);
  console.log(`  headful     ${HEADFUL}`);
  console.log(`  inspect     ${INSPECT}`);
  console.log(`  debug       ${DEBUG}`);
  console.log('====================');

  const isFirstRun = !existsSync(path.join(AUTH_DIR, 'Default'));
  dbg('launching persistent chromium context', { headless: !HEADFUL, isFirstRun });
  dbg('auth dir exists =', existsSync(AUTH_DIR), ', "Default" profile exists =', existsSync(path.join(AUTH_DIR, 'Default')));

  // Enterprise Teams (and Entra conditional access) commonly invalidates sessions
  // when it detects automation. These flags make Playwright's Chromium look like
  // a normal Chrome session: no "Chrome is being controlled by automated test
  // software" banner, no navigator.webdriver flag, realistic UA.
  const launchOpts = {
    headless: !HEADFUL,
    viewport: { width: 1400, height: 1000 },
    args: [
      '--disable-blink-features=AutomationControlled',
      '--disable-features=Translate,IsolateOrigins,site-per-process',
    ],
    ignoreDefaultArgs: ['--enable-automation'],
    userAgent:
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
      '(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
  };
  // If the user has real Chrome installed, prefer it — profiles created by real
  // Chrome survive Entra device-binding checks better than bundled Chromium.
  if (args.channel) launchOpts.channel = String(args.channel); // e.g. --channel chrome

  const context = await chromium.launchPersistentContext(AUTH_DIR, launchOpts);

  // Persist the session on Ctrl+C — otherwise the profile can be left in a
  // half-written state and next run needs to re-auth.
  let closing = false;
  const gracefulClose = async (sig) => {
    if (closing) return;
    closing = true;
    console.log(`\n→ Caught ${sig}, closing browser to save session...`);
    try { await context.close(); } catch {}
    process.exit(130);
  };
  process.on('SIGINT', () => gracefulClose('SIGINT'));
  process.on('SIGTERM', () => gracefulClose('SIGTERM'));

  const page = context.pages()[0] ?? await context.newPage();

  page.on('console', msg => {
    if (DEBUG && ['error', 'warning'].includes(msg.type())) {
      console.log(`[page.${msg.type()}]`, msg.text().slice(0, 240));
    }
  });

  console.log(`→ Navigating to ${TEAMS_URL}`);
  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });
  dbg('post-goto url =', page.url());

  console.log('→ Waiting for Teams shell (up to 5 minutes; sign in if prompted)...');

  // Wait for a Teams-specific element — NOT [role="main"], which also matches
  // the login page and causes us to plow ahead before sign-in finishes.
  try {
    await page.waitForFunction(() => {
      if (/login\.microsoftonline\.com|login\.live\.com/i.test(location.hostname)) return false;
      return !!document.querySelector(
        '[data-tid="app-bar-wrapper"], [data-tid="chat-pane"], [data-tid="message-pane-list-surface"], [data-tid="searchBoxInput"], [data-tid="topBarSearchInput"]'
      );
    }, { timeout: 5 * 60_000, polling: 1000 });
  } catch {
    const url = page.url();
    console.error('✗ Timed out waiting for Teams shell.');
    console.error(`  Final URL: ${url}`);
    if (/login\./i.test(url)) {
      console.error('  Still on a Microsoft login page — sign in completed? Try again, and complete MFA if prompted.');
    } else {
      console.error('  Not on a known login page; tenant may require conditional access / device compliance.');
    }
    await context.close();
    process.exit(3);
  }
  await page.waitForTimeout(1500);
  console.log('✓ Teams shell detected.');
  dbg('shell url =', page.url(), ', title =', await page.title());

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
      const opened = await gotoTarget(page, t);
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

async function openSearchBox(page) {
  // Try clicking the top search box first (most reliable in Teams v2).
  const result = await page.evaluate(() => {
    const sels = [
      '[data-tid="searchBoxInput"]',
      '[data-tid="search-box"]',
      '[data-tid="topBarSearchInput"]',
      '[data-tid="searchV2SearchBox"]',
      '[data-tid*="search" i] input',
      '[data-tid*="search" i] [contenteditable="true"]',
      '[placeholder*="Search" i]',
      'input[aria-label*="Search" i]',
      '[aria-label*="Search" i][contenteditable="true"]',
      '[role="search"] input',
      '[role="search"] [contenteditable="true"]',
      // Some tenants expose a Search button that opens the real input.
      'button[aria-label*="Search" i]',
    ];
    for (const sel of sels) {
      const els = document.querySelectorAll(sel);
      for (const el of els) {
        const rect = el.getBoundingClientRect();
        if (rect.width > 40 && rect.height > 10) {
          el.focus();
          el.click?.();
          return { clicked: true, via: sel, tag: el.tagName, rect: { w: rect.width, h: rect.height } };
        }
      }
    }
    // Diagnostic dump: list any visible element whose text/aria mentions "search".
    const visible = [];
    for (const el of document.querySelectorAll('input, button, [contenteditable="true"], [role="search"]')) {
      const rect = el.getBoundingClientRect();
      if (rect.width < 40 || rect.height < 10) continue;
      const aria = (el.getAttribute('aria-label') || '').toLowerCase();
      const ph = (el.getAttribute('placeholder') || '').toLowerCase();
      const tid = el.getAttribute('data-tid') || '';
      if (aria.includes('search') || ph.includes('search') || /search/i.test(tid)) {
        visible.push({ tag: el.tagName, tid, aria, ph, w: Math.round(rect.width), h: Math.round(rect.height) });
      }
    }
    return {
      clicked: false,
      counts: Object.fromEntries(sels.map(s => [s, document.querySelectorAll(s).length])),
      visibleSearchLike: visible.slice(0, 8),
    };
  });
  if (result.clicked) {
    dbg(`openSearchBox: clicked via ${result.via} (${result.tag}, ${result.rect.w}x${result.rect.h})`);
    return true;
  }
  dbg('openSearchBox: no direct match, counts =', JSON.stringify(result.counts));
  dbg('openSearchBox: visible search-like elements =', JSON.stringify(result.visibleSearchLike));

  // Keyboard fallbacks.
  dbg('openSearchBox: trying Ctrl+E');
  await page.keyboard.press('Control+E').catch(() => {});
  await page.waitForTimeout(400);
  const afterCtrlE = await page.evaluate(() => ({
    active: document.activeElement?.tagName,
    activeRole: document.activeElement?.getAttribute('role'),
    activePlaceholder: document.activeElement?.getAttribute('placeholder'),
  }));
  dbg('openSearchBox: after Ctrl+E, active =', JSON.stringify(afterCtrlE));
  await page.keyboard.press('Control+Alt+E').catch(() => {});
  await page.waitForTimeout(400);
  return true;
}

async function gotoTarget(page, target) {
  const spec = (typeof target === 'string') ? { name: target } : target;
  const displayName = spec.label || spec.name || `${spec.team || ''} / ${spec.channel || ''}`.trim();
  const beforeHeader = await page.evaluate(readHeader).catch(() => ({ header: null }));
  dbg(`gotoTarget("${displayName}"): before header = ${beforeHeader.header ?? '(none)'}`);

  // Dismiss any search overlay left over from a prior target, and make sure
  // the Chat tab is active so the left rail is actually rendered.
  await page.keyboard.press('Escape').catch(() => {});
  const tab = await page.evaluate(ensureChatTabActive).catch(() => null);
  if (tab?.clicked) { dbg(`gotoTarget: activated Chat tab via ${tab.via}`); await page.waitForTimeout(600); }

  // Try sidebar tree first, with up to a few scroll passes to render
  // virtualized items.
  const sidebarSpec = { name: spec.name, team: spec.team, channel: spec.channel };
  let sidebar = null;
  for (let attempt = 0; attempt < 6; attempt++) {
    sidebar = await page.evaluate(clickSidebarTarget, sidebarSpec).catch((e) => ({ error: String(e) }));
    if (sidebar?.clicked) break;
    // No match yet — scroll the rail and retry.
    const mv = await page.evaluate(scrollSidebarTree, 700).catch(() => null);
    dbg(`gotoTarget: sidebar miss on attempt ${attempt}, scroll=${mv ? JSON.stringify({ before: mv.before, after: mv.after, atBottom: mv.atBottom }) : 'n/a'}`);
    if (!mv?.ok || mv.atBottom || mv.before === mv.after) break;
    await page.waitForTimeout(300);
  }

  if (sidebar?.clicked) {
    const label = sidebar.mode === 'team-channel'
      ? `team="${sidebar.teamMatch}" channel="${sidebar.channelMatch}" (teamScore=${sidebar.teamScore}, chScore=${sidebar.channelScore})`
      : `"${sidebar.matchText}" (score=${sidebar.score})`;
    dbg(`gotoTarget: sidebar hit ${label}, expanded=${sidebar.expandedFolders}`);
    const beforeUrl = page.url();
    const ok = await page.waitForFunction(
      ({ prev, prevUrl }) => {
        // URL changed? Channels and chats both swap the URL on navigation.
        if (location.href !== prevUrl) return true;
        const hEl = document.querySelector(
          '[data-tid="chat-header-title"], [data-tid="chat-title"], [data-tid="channel-header-title"], [data-tid="channel-header"], [role="main"] h1, [role="main"] h2'
        );
        const h = hEl?.textContent?.trim() || null;
        if (h && h !== prev) return true;
        // Any message/post/reply renderer visible?
        const hasMsgs = document.querySelectorAll(
          '[data-tid="chat-pane-item"], [data-tid="chat-pane-message"], [data-tid="message-pane-item"], [data-tid^="post-message-renderer"], [data-tid^="reply-message-renderer"], [data-tid="message-pane-list-surface"], [data-tid="message-pane-list"], [data-tid="threadBodyList"], [data-tid="channel-content"]'
        ).length > 0;
        return hasMsgs;
      },
      { prev: beforeHeader.header, prevUrl: beforeUrl },
      { timeout: 15_000 }
    ).then(() => true).catch(() => false);
    const afterHeader = await page.evaluate(readHeader).catch(() => ({ header: null }));
    dbg(`gotoTarget: sidebar success=${ok}, after header = ${afterHeader.header ?? '(none)'}`);
    if (ok) return true;
    dbg('gotoTarget: sidebar click did not settle, falling back to search');
  } else {
    const m = sidebar?.mode || 'name';
    if (m === 'team-channel') {
      dbg(`gotoTarget: sidebar miss (team-channel, reason=${sidebar?.reason})`);
      if (sidebar?.teamCandidates?.length) {
        dbg(`  team candidates: ${sidebar.teamCandidates.map(c => `"${c.text}"(${c.score})`).join(', ')}`);
      }
      if (sidebar?.channelCandidates?.length) {
        dbg(`  channel candidates under "${sidebar.teamMatch}": ${sidebar.channelCandidates.map(c => `"${c.text}"(${c.score})`).join(', ')}`);
      }
    } else {
      dbg(`gotoTarget: sidebar miss (candidates=${sidebar?.candidateCount ?? 0}, expanded=${sidebar?.expandedFolders ?? 0})`);
      if (sidebar?.topCandidates?.length) {
        dbg(`  top candidates: ${sidebar.topCandidates.map(c => `"${c.text}"(${c.score}${c.kind ? `,${c.kind}` : ''})`).join(', ')}`);
      }
    }
  }

  await openSearchBox(page);
  await page.waitForTimeout(500);

  const opened = await pickFromSuggestions(page, spec.name);
  if (!opened) {
    dbg('gotoTarget: pickFromSuggestions returned false');
    return false;
  }

  const success = await page.waitForFunction(
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

  const afterHeader = await page.evaluate(readHeader).catch(() => ({ header: null }));
  dbg(`gotoTarget: search success=${success}, after header = ${afterHeader.header ?? '(none)'}`);
  return success;
}

async function pickFromSuggestions(page, name) {
  // Use whatever is focused (the search box we just opened).
  const activeSel = 'input[type="text"]:focus, input[type="search"]:focus, [role="combobox"]:focus, [role="searchbox"]:focus, [contenteditable="true"]:focus';
  let input = await page.$(activeSel);
  let via = 'focused';
  if (!input) {
    input = await page.$('[data-tid="searchBoxInput"], [data-tid="topBarSearchInput"], [placeholder*="Search" i], [role="search"] input, [role="search"] [contenteditable="true"]');
    via = 'fallback-selector';
  }
  if (!input) {
    dbg('pickFromSuggestions: no input element found');
    return false;
  }
  dbg(`pickFromSuggestions: using input via ${via}`);

  await input.click({ delay: 30 }).catch(() => {});
  await page.keyboard.press('Control+A').catch(() => {});
  await page.keyboard.press('Delete').catch(() => {});
  await page.keyboard.type(name, { delay: 30 });
  await page.waitForTimeout(1200);

  const result = await page.evaluate((needle) => {
    const lower = String(needle).toLowerCase();
    const tokens = lower.split(/\s+/).filter(Boolean);
    const sels = [
      '[role="option"]',
      '[role="listitem"]',
      '[data-tid^="suggestion"]',
      '[data-tid^="search-result"]',
      '[data-tid^="searchResult"]',
      '[data-tid*="chat-list-item"]',
      '[data-tid*="team-channel"]',
    ];
    const candidates = document.querySelectorAll(sels.join(','));
    const counts = Object.fromEntries(sels.map(s => [s, document.querySelectorAll(s).length]));
    const samples = [...candidates].slice(0, 5).map(el => (el.textContent || '').trim().slice(0, 80));
    // Prefer exact/contiguous matches.
    let best = null, bestScore = -1, bestText = null;
    for (const el of candidates) {
      const t = (el.textContent || '').toLowerCase();
      if (!t) continue;
      let score = 0;
      if (t.includes(lower)) score = 100;
      else if (tokens.every(tok => t.includes(tok))) score = 50;
      else continue;
      score -= Math.min(20, Math.floor(t.length / 200));
      if (score > bestScore) { best = el; bestScore = score; bestText = t.slice(0, 80); }
    }
    if (best) { best.click(); return { clicked: true, score: bestScore, text: bestText, total: candidates.length, counts, samples }; }
    return { clicked: false, total: candidates.length, counts, samples };
  }, name);

  dbg(`pickFromSuggestions: ${result.total} candidates, counts=${JSON.stringify(result.counts)}`);
  if (result.samples?.length) dbg('pickFromSuggestions: samples =', result.samples);
  if (result.clicked) {
    dbg(`pickFromSuggestions: clicked score=${result.score} text="${result.text}"`);
  } else {
    dbg('pickFromSuggestions: no match — pressing Enter');
    await page.keyboard.press('Enter').catch(() => {});
  }
  await page.waitForTimeout(800);
  return true;
}

async function harvestOpenPane(page, cutoff) {
  const seen = new Map();
  let reachedOlder = false;
  let stagnant = 0;
  let lastTop = -1;

  for (let round = 0; round < MAX_SCROLL_ROUNDS && !reachedOlder && seen.size < MAX_MESSAGES && stagnant < 3; round++) {
    const exp = await page.evaluate(expandChannelReplies).catch(() => ({ clicked: 0 }));
    if (exp?.clicked) {
      dbg(`harvest round ${round}: expanded ${exp.clicked} reply/see-more button(s)`);
      await page.waitForTimeout(400);
    }
    const snap = await page.evaluate(collectPaneMessages);
    dbg(`harvest round ${round}: +${snap.msgs.length} raw msgs, hasContainer=${snap.hasContainer}, hasScroller=${snap.hasScroller}, scroll=${JSON.stringify(snap.scroll)}`);
    for (const m of snap.msgs) {
      const key = `${m.timeISO || ''}|${m.sender || ''}|${(m.text || '').slice(0, 80)}`;
      if (seen.has(key)) continue;
      seen.set(key, m);
      if (m.timeISO && new Date(m.timeISO) < cutoff) reachedOlder = true;
    }
    if (reachedOlder || seen.size >= MAX_MESSAGES) break;

    const mv = await page.evaluate(scrollPaneUpBy, 1400);
    if (!mv?.ok) { dbg('harvest: no scrollable container, stopping'); break; }
    if (mv.atTop || mv.after === lastTop) { stagnant++; dbg(`harvest: stagnant=${stagnant} (atTop=${mv.atTop}, top=${mv.after})`); }
    else { stagnant = 0; lastTop = mv.after; }
    await page.waitForTimeout(PANE_SETTLE_MS);
  }

  const header = (await page.evaluate(readHeader)).header;
  const messages = [...seen.values()]
    .filter(m => !m.timeISO || new Date(m.timeISO) >= cutoff)
    .sort((a, b) => new Date(a.timeISO || 0) - new Date(b.timeISO || 0));
  dbg(`harvest done: ${seen.size} unique, ${messages.length} after cutoff filter, reachedOlder=${reachedOlder}`);
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

// DOM helpers for enterprise Microsoft Teams (teams.microsoft.com/v2).
//
// Each exported function is fully self-contained: selector banks + helpers are
// inlined inside each function body. Reason: Playwright's `page.evaluate(fn)`
// serializes only that one function's source, so module-scope constants and
// helper functions are NOT available in the browser context.

export function readHeader() {
  const hEl = document.querySelector(
    '[data-tid="chat-header-title"], [data-tid="chat-title"], [data-tid="channel-header-title"], [data-tid="channelTitle-text"], [data-tid="channel-header"] h1, [data-tid="channel-header"] h2, [role="banner"] h1, [role="main"] h1, [role="main"] h2'
  );
  return { header: hEl?.textContent?.trim() || null, title: document.title };
}

export function collectPaneMessages() {
  const CHAT_CONTAINER_SELECTORS = [
    '.fui-Chat',
    '[class*="fui-Chat"]',
    '[data-tid="message-pane-list-surface"]',
    '[data-tid="message-pane-list"]',
    '[data-tid="channel-pane-viewport"]',
    '[data-tid="channel-content"]',
    '[data-tid="channel-pane-runway"]',
    '[id="channel-pane"]',
    '[data-tid="threadBodyList"]',
    '.ts-message-list-container',
    '[data-tid="chat-pane"]',
    '[data-shortcut-context="chat-messages-list"]',
    '[role="main"] [data-is-scrollable="true"]',
    '.message-list',
    '[class*="message-list"]',
    '[class*="MessageList"]',
    '[data-tid="messageListContainer"]',
    '[role="log"]',
  ];
  const MESSAGE_SELECTORS = [
    '[data-tid="chat-pane-item"]',
    '[data-tid="message-pane-item"]',
    '[id^="post-message-renderer-"]',
    '[id^="reply-message-renderer-"]',
    '[id^="message-body-"]',
    '[data-testid="message-wrapper"]',
    '[data-tid="chat-pane-message"]',
    '[class*="fui-unstable-ChatItem"]',
    '[data-tid="messageWrapper"]',
    '.message-body-container',
    '[class*="message-item"]',
    '[role="listitem"]',
  ];
  const TIMESTAMP_SELECTORS = [
    '[class*="fui-ChatMessage__timestamp"]',
    '[class*="fui-ChatMyMessage__timestamp"]',
    '[class*="__timestamp"]',
    '[data-tid="messageTimeStamp"]',
    'time[datetime]',
    'time',
    '[class*="timestamp"]',
    '[class*="Timestamp"]',
    '[class*="time-stamp"]',
    '[datetime]',
    '[data-tid^="timestamp-"]',
  ];
  const SENDER_SELECTORS = [
    '[class*="fui-ChatMessage__author"]',
    '[class*="fui-ChatMyMessage__author"]',
    '[class*="__author"]',
    '[data-tid="message-author-name"]',
    '.ui-chat__message__author',
    '[class*="author"]',
    '[class*="Author"]',
    '[class*="sender"]',
    '[class*="Sender"]',
    '[class*="displayName"]',
    '[class*="DisplayName"]',
    '[data-tid^="author-"]',
  ];
  const MESSAGE_TEXT_SELECTORS = [
    '[class*="fui-ChatMessage__body"]',
    '[class*="fui-ChatMyMessage__body"]',
    '[class*="__body"]',
    '.fui-ChatMessage',
    '.fui-ChatMyMessage',
    '.ui-chat__message__content',
    '.message-body-content',
    '[class*="messageContent"]',
    '[class*="MessageContent"]',
    '[class*="message-body"]',
    '[data-tid="messageBodyContent"]',
    '[data-tid="message-body-content"]',
    '[data-tid="message-body"]',
  ];

  const q = (selectors, parent) => {
    const root = parent || document;
    for (const sel of selectors) {
      try { const el = root.querySelector(sel); if (el) return el; } catch {}
    }
    return null;
  };
  const qa = (selectors, parent) => {
    const root = parent || document;
    for (const sel of selectors) {
      try { const els = root.querySelectorAll(sel); if (els.length) return [...els]; } catch {}
    }
    return [];
  };
  const getChatContainer = () => {
    for (const sel of CHAT_CONTAINER_SELECTORS) {
      try {
        for (const el of document.querySelectorAll(sel)) {
          if ((el.textContent || '').trim().length > 80 && el.children.length > 0) return el;
        }
      } catch {}
    }
    const probes = document.querySelectorAll(
      '[role="main"], [role="log"], [role="list"], [data-tid="chat"], [class*="chat"], [class*="Chat"], [class*="message-list"]'
    );
    for (const el of probes) {
      if (el.children.length >= 3 && (el.textContent || '').length > 200) return el;
    }
    return null;
  };
  const getScrollableContainer = () => {
    const chat = getChatContainer();
    if (chat) {
      if (chat.scrollHeight > chat.clientHeight + 10) return chat;
      let p = chat.parentElement;
      while (p) {
        const cs = getComputedStyle(p);
        if ((cs.overflowY === 'auto' || cs.overflowY === 'scroll') && p.scrollHeight > p.clientHeight + 10) return p;
        p = p.parentElement;
      }
      const child = chat.querySelector('[style*="overflow"], [data-is-scrollable="true"]');
      if (child && child.scrollHeight > child.clientHeight + 10) return child;
    }
    const all = document.querySelectorAll('[role="main"] *, [class*="chat"] *, [class*="Chat"] *');
    let best = null;
    for (const el of all) {
      if (el.scrollHeight > el.clientHeight + 100) {
        const cs = getComputedStyle(el);
        if (cs.overflowY === 'auto' || cs.overflowY === 'scroll') {
          if (!best || el.scrollHeight > best.scrollHeight) best = el;
        }
      }
    }
    return best || chat;
  };
  const parseTimestamp = (raw) => {
    if (!raw) return null;
    const s = String(raw).trim();
    const now = new Date();
    const direct = new Date(s);
    if (!isNaN(direct.getTime())) return direct;
    if (/^today/i.test(s)) {
      const t = s.replace(/today\s*,?\s*/i, '');
      const d = new Date(`${now.toDateString()} ${t}`);
      if (!isNaN(d.getTime())) return d;
    }
    if (/^yesterday/i.test(s)) {
      const t = s.replace(/yesterday\s*,?\s*/i, '');
      const y = new Date(now); y.setDate(y.getDate() - 1);
      const d = new Date(`${y.toDateString()} ${t}`);
      if (!isNaN(d.getTime())) return d;
    }
    if (/^\d{1,2}:\d{2}\s*(AM|PM)?$/i.test(s)) {
      const d = new Date(`${now.toDateString()} ${s}`);
      if (!isNaN(d.getTime())) return d;
    }
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*(.*)$/);
    if (m) {
      const [, mo, da, yr, t] = m;
      const year = yr.length === 2 ? `20${yr}` : yr;
      const d = new Date(`${mo}/${da}/${year} ${t || '00:00'}`);
      if (!isNaN(d.getTime())) return d;
    }
    return null;
  };

  const container = getChatContainer();
  if (!container) return { header: null, msgs: [], scroll: null, hasContainer: false, hasScroller: false };

  const messageEls = qa(MESSAGE_SELECTORS, container);
  let lastSender = 'Unknown';
  let lastTime = null;
  let lastTimeStr = '';
  const msgs = [];

  messageEls.forEach((el, idx) => {
    const item = el.closest('[data-tid="chat-pane-item"]')
      || el.closest('[data-tid="chat-pane-message"]')
      || el.closest('[data-tid="message-pane-item"]')
      || el.closest('[id^="post-message-renderer-"]')
      || el.closest('[id^="reply-message-renderer-"]')
      || el;

    if (item.querySelector('[class*="fui-Divider"]')
      || /fui-Divider/.test(String(item.className || ''))) return;

    const classStr = String(item.className || '');
    const isControl = !!(item.querySelector('[class*="ChatControlMessage"]')
      || item.querySelector('[class*="ControlMessage"]')
      || /ChatControlMessage|ControlMessage/.test(classStr));

    if (isControl) {
      const controlEl = item.querySelector('[class*="ChatControlMessage"] [role="heading"]')
        || item.querySelector('[class*="ControlMessage"]') || item;
      const text = controlEl.textContent?.trim() || '';
      if (text) {
        msgs.push({
          isSystemMessage: true, sender: '[System]',
          text: text.slice(0, 2000),
          timeISO: lastTime ? lastTime.toISOString() : null,
          timeLabel: lastTimeStr || 'Chat start',
        });
      }
      return;
    }

    let timeStr = '';
    for (const sel of TIMESTAMP_SELECTORS) {
      try {
        const found = item.querySelector(sel);
        if (found) {
          timeStr = found.getAttribute('datetime') || found.getAttribute('title') || found.textContent?.trim() || '';
          if (timeStr) break;
        }
      } catch {}
    }
    if (!timeStr) {
      for (const span of item.querySelectorAll('span, div, time')) {
        const t = span.textContent?.trim() || '';
        if (/^\d{1,2}\/\d{1,2}\/\d{2,4}\s+\d{1,2}:\d{2}\s*(AM|PM)/i.test(t)
          || /^\d{1,2}:\d{2}\s*(AM|PM)/i.test(t)) { timeStr = t; break; }
      }
    }
    const parsedTime = parseTimestamp(timeStr);
    if (parsedTime) { lastTime = parsedTime; lastTimeStr = timeStr; }

    let sender = '';
    const senderEl = q(SENDER_SELECTORS, item);
    if (senderEl) sender = senderEl.textContent?.trim() || '';
    if (!sender) {
      for (const span of item.querySelectorAll('span, div')) {
        const t = span.textContent?.trim() || '';
        if (/^[A-Z][a-z]+,\s*[A-Z][a-z]+(\s*\([A-Z]+\))?$/.test(t)) { sender = t; break; }
      }
    }
    if (sender) lastSender = sender;

    let text = '';
    const bodyEl = q(MESSAGE_TEXT_SELECTORS, el);
    if (bodyEl) text = bodyEl.textContent?.trim() || '';
    if (!text) text = el.textContent?.trim() || '';
    if (text && sender) {
      const esc = sender.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      text = text.replace(new RegExp(`^.*?by\\s+${esc}`, 'i'), '').trim();
      text = text.replace(new RegExp(`^${esc}`), '').trim();
    }
    if (text && timeStr) text = text.replace(timeStr, '').trim();
    if (!text) return;

    msgs.push({
      isSystemMessage: false,
      sender: sender || lastSender,
      text: text.slice(0, 5000),
      timeISO: (parsedTime || lastTime)?.toISOString() || null,
      timeLabel: timeStr || lastTimeStr || null,
    });
  });

  const scroller = getScrollableContainer();
  const hEl = document.querySelector(
    '[data-tid="chat-header-title"], [data-tid="chat-title"], [data-tid="channel-header-title"], [role="banner"] h1, [role="main"] h1, [role="main"] h2'
  );
  return {
    header: hEl?.textContent?.trim() || null,
    msgs,
    scroll: scroller
      ? { scrollTop: scroller.scrollTop, scrollHeight: scroller.scrollHeight, clientHeight: scroller.clientHeight }
      : null,
    hasContainer: true,
    hasScroller: !!scroller,
  };
}

export function scrollPaneUpBy(delta) {
  const CHAT_CONTAINER_SELECTORS = [
    '.fui-Chat', '[class*="fui-Chat"]', '[data-tid="message-pane-list-surface"]',
    '[data-tid="message-pane-list"]', '[data-tid="channel-pane-viewport"]',
    '[data-tid="channel-content"]',
    '[data-tid="channel-pane-runway"]', '[id="channel-pane"]',
    '[data-tid="threadBodyList"]',
    '.ts-message-list-container', '[data-tid="chat-pane"]',
    '[data-shortcut-context="chat-messages-list"]',
    '[role="main"] [data-is-scrollable="true"]', '.message-list',
    '[class*="message-list"]', '[class*="MessageList"]',
    '[data-tid="messageListContainer"]', '[role="log"]',
  ];
  const getChatContainer = () => {
    for (const sel of CHAT_CONTAINER_SELECTORS) {
      try {
        for (const el of document.querySelectorAll(sel)) {
          if ((el.textContent || '').trim().length > 80 && el.children.length > 0) return el;
        }
      } catch {}
    }
    return null;
  };
  const getScroller = () => {
    const isScrollable = (node) => {
      if (!node || node === document.body) return false;
      const cs = window.getComputedStyle(node);
      return (cs.overflowY === 'auto' || cs.overflowY === 'scroll')
        && node.scrollHeight > node.clientHeight + 10;
    };
    const chat = getChatContainer();
    if (chat) {
      if (isScrollable(chat)) return chat;
      let p = chat.parentElement;
      while (p && p !== document.documentElement) {
        if (isScrollable(p)) return p;
        p = p.parentElement;
      }
    }
    // Anchor off a known post element for channels.
    for (const anchor of document.querySelectorAll('[id^="post-message-renderer-"], [id^="message-body-"], [data-tid="message-pane-item"]')) {
      let p = anchor.parentElement;
      while (p && p !== document.documentElement) {
        if (isScrollable(p)) return p;
        p = p.parentElement;
      }
    }
    // Last resort: biggest scrollable inside [role="main"].
    const main = document.querySelector('[role="main"]');
    if (main && isScrollable(main)) return main;
    if (main) {
      let best = null;
      for (const node of main.querySelectorAll('*')) {
        if (isScrollable(node) && (!best || node.scrollHeight > best.scrollHeight)) best = node;
      }
      if (best) return best;
    }
    return chat;
  };
  const el = getScroller();
  if (!el) return { ok: false };
  const before = el.scrollTop;
  el.scrollTop = Math.max(0, el.scrollTop - Number(delta || 800));
  return { ok: true, before, after: el.scrollTop, atTop: el.scrollTop <= 1, scrollHeight: el.scrollHeight };
}

export function scrollPaneToTop() {
  const sels = [
    '.fui-Chat', '[class*="fui-Chat"]', '[data-tid="message-pane-list-surface"]',
    '[data-tid="message-pane-list"]', '[data-tid="channel-pane-viewport"]',
    '[data-tid="channel-content"]',
    '[data-tid="channel-pane-runway"]', '[id="channel-pane"]',
    '[data-tid="chat-pane"]', '[data-shortcut-context="chat-messages-list"]',
    '[role="main"] [data-is-scrollable="true"]',
    '[data-tid="messageListContainer"]', '[role="log"]',
  ];
  let el = null;
  for (const sel of sels) {
    const found = document.querySelector(sel);
    if (found && found.scrollHeight > found.clientHeight + 10) { el = found; break; }
  }
  if (!el) return { ok: false };
  el.scrollTop = 0;
  return { ok: true };
}

// Scroll the message pane all the way to the bottom. Teams opens chats/channels
// at the last-read position, so unread messages below it are missed unless we
// jump to the newest first and then harvest backwards.
export function scrollPaneToBottom() {
  const ANCHOR_SELS = [
    '[data-tid="channel-pane-viewport"]',
    '.fui-Chat', '[class*="fui-Chat"]', '[data-tid="message-pane-list-surface"]',
    '[data-tid="message-pane-list"]', '[data-tid="channel-content"]',
    '[data-tid="channel-pane-runway"]', '[id="channel-pane"]',
    '[data-tid="chat-pane"]', '[data-shortcut-context="chat-messages-list"]',
    '[role="main"] [data-is-scrollable="true"]',
    '[data-tid="messageListContainer"]', '[role="log"]',
    // Posts anchor — walk up from a known post element to find scroller.
    '[id^="post-message-renderer-"]', '[id^="message-body-"]',
    '[data-tid="chat-pane-item"]', '[data-tid="message-pane-item"]',
  ];
  const isScrollable = (el) => {
    if (!el || el === document.body) return false;
    const cs = window.getComputedStyle(el);
    return (cs.overflowY === 'auto' || cs.overflowY === 'scroll')
      && el.scrollHeight > el.clientHeight + 10;
  };
  // 1) Try direct selection.
  let el = null;
  for (const sel of ANCHOR_SELS) {
    try {
      const found = document.querySelector(sel);
      if (!found) continue;
      if (isScrollable(found)) { el = found; break; }
      // Walk up to find scrollable ancestor.
      let p = found.parentElement;
      while (p && p !== document.documentElement) {
        if (isScrollable(p)) { el = p; break; }
        p = p.parentElement;
      }
      if (el) break;
    } catch {}
  }
  // 2) Last resort: biggest scrollable element inside [role="main"].
  if (!el) {
    const main = document.querySelector('[role="main"]');
    if (main && isScrollable(main)) { el = main; }
    if (!el && main) {
      let best = null;
      for (const node of main.querySelectorAll('*')) {
        if (isScrollable(node)) {
          if (!best || node.scrollHeight > best.scrollHeight) best = node;
        }
      }
      el = best;
    }
  }
  if (!el) return { ok: false, reason: 'no-scroller' };
  const before = el.scrollTop;
  el.scrollTop = el.scrollHeight;
  return {
    ok: true, before, after: el.scrollTop,
    atBottom: el.scrollTop + el.clientHeight >= el.scrollHeight - 2,
    scrollHeight: el.scrollHeight,
  };
}

export function inspectPane() {
  const CHAT_CONTAINER_SELECTORS = [
    '.fui-Chat', '[class*="fui-Chat"]', '[data-tid="message-pane-list-surface"]',
    '[data-tid="message-pane-list"]', '[data-tid="channel-pane-viewport"]',
    '[data-tid="channel-content"]',
    '[data-tid="channel-pane-runway"]', '[id="channel-pane"]',
    '[data-tid="threadBodyList"]',
    '.ts-message-list-container', '[data-tid="chat-pane"]',
    '[data-shortcut-context="chat-messages-list"]',
    '[role="main"] [data-is-scrollable="true"]', '.message-list',
    '[class*="message-list"]', '[class*="MessageList"]',
    '[data-tid="messageListContainer"]', '[role="log"]',
  ];
  const MESSAGE_SELECTORS = [
    '[data-tid="chat-pane-item"]', '[data-tid="message-pane-item"]',
    '[id^="post-message-renderer-"]', '[id^="reply-message-renderer-"]',
    '[id^="message-body-"]',
    '[data-testid="message-wrapper"]',
    '[data-tid="chat-pane-message"]', '[class*="fui-unstable-ChatItem"]',
    '[data-tid="messageWrapper"]', '.message-body-container',
    '[class*="message-item"]', '[role="listitem"]',
  ];
  let chat = null;
  for (const sel of CHAT_CONTAINER_SELECTORS) {
    try {
      for (const el of document.querySelectorAll(sel)) {
        if ((el.textContent || '').trim().length > 80 && el.children.length > 0) { chat = el; break; }
      }
      if (chat) break;
    } catch {}
  }
  const out = {
    url: location.href, title: document.title,
    containerFound: !!chat,
    containerTag: chat?.tagName || null,
    containerClass: chat?.className?.toString().slice(0, 120) || null,
    selectorCounts: {},
  };
  for (const sel of MESSAGE_SELECTORS) {
    try { out.selectorCounts[sel] = document.querySelectorAll(sel).length; } catch {}
  }
  return out;
}

// Channels lazy-render replies: each post shows up to ~3 recent replies behind
// an "Open N replies from ..." button. This expands all such buttons in the
// visible DOM so collectPaneMessages can see the reply bodies. Also expands
// "See more" truncation buttons inside long message bodies.
export function expandChannelReplies() {
  const patterns = [
    /^\s*(open|show|see|view)\s+\d+\s+(older\s+)?repl(y|ies)/i,
    /^\s*\d+\s+(older\s+)?repl(y|ies)/i,
    /hidden\s+repl(y|ies)/i,
    /^\s*see\s+more\b/i,
  ];
  const nodes = document.querySelectorAll('button, [role="button"]');
  let clicked = 0;
  for (const btn of nodes) {
    if (btn.getAttribute('aria-disabled') === 'true' || btn.disabled) continue;
    const label = (btn.getAttribute('aria-label') || btn.textContent || '').trim();
    if (!label) continue;
    if (patterns.some((re) => re.test(label))) {
      try { btn.click(); clicked++; } catch {}
    }
  }
  return { clicked };
}

// Navigate via the left sidebar tree (Chat list / Teams list) instead of search.
// Accepts either a bare name string, or an object: { name?, team?, channel? }.
// When { team, channel } is provided, finds the team treeitem first, expands
// it if collapsed, then scopes the channel search to that team's subtree.
// Returns a diagnostics object describing what was found/clicked.
export function clickSidebarTarget(target) {
  const spec = (typeof target === 'string') ? { name: target } : (target || {});
  const needleName = String(spec.name || '').trim().toLowerCase();
  const needleTeam = String(spec.team || '').trim().toLowerCase();
  const needleChannel = String(spec.channel || '').trim().toLowerCase();

  const norm = (s) => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
  // Very light similarity: ratio of matching chars by sliding window; cheap,
  // tolerates 1-2 character typos (e.g. "Shapshifters" vs "ShapeShifters").
  const similarity = (a, b) => {
    if (!a || !b) return 0;
    const shorter = a.length < b.length ? a : b;
    const longer = a.length < b.length ? b : a;
    if (longer.length === 0) return 1;
    let matches = 0;
    for (let i = 0; i < shorter.length; i++) {
      if (longer.includes(shorter.slice(i, i + 3))) matches++;
    }
    return matches / shorter.length;
  };
  const score = (title, needle) => {
    const t = norm(title);
    const n = norm(needle);
    if (!t || !n) return 0;
    if (t === n) return 100;
    if (t.startsWith(n)) return 80;
    const words = n.split(/\s+/).filter(Boolean);
    if (words.length > 1 && words.every((w) => t.includes(w))) return 50 + Math.min(20, words.length * 5);
    if (t.includes(n)) return 40;
    // fuzzy fallback: close enough for typos
    const sim = similarity(t, n);
    if (sim >= 0.8) return 35;
    if (sim >= 0.65) return 25;
    return 0;
  };

  const expandFolder = (folder) => {
    if (!folder) return false;
    if (folder.getAttribute('aria-expanded') === 'true') return false;
    const header = folder.querySelector('[data-testid="conversation-folder-header"], [data-inp="simple-collab-folder-header"]') || folder;
    try { header.click(); return true; } catch { return false; }
  };

  // 1) Expand top-level folders (Chat list, Teams & channels, Favorites, etc.)
  let expanded = 0;
  const topFolders = document.querySelectorAll(
    '[data-testid="simple-collab-dnd-rail"] [data-item-type="custom-folder"][aria-expanded="false"], '
    + '[data-testid="simple-collab-dnd-rail"] [data-item-type="teams-and-channels"][aria-expanded="false"], '
    + '[data-testid="simple-collab-dnd-rail"] [data-item-type="chats"][aria-expanded="false"]'
  );
  for (const f of topFolders) if (expandFolder(f)) expanded++;

  // Helper: click a treeitem (prefer its main switch/content area).
  const clickTreeitem = (el) => {
    const item = el.closest('[role="treeitem"]') || el;
    const clickable = item.querySelector('[data-inp="simple-collab-chat-switch"], [data-inp="simple-collab-channel-switch"], [data-inp="simple-collab-list-item-teams-and-channels"]')
      || item.querySelector('[data-testid="list-item"]')
      || item.querySelector('[class*="TreeItemLayout__main"]')
      || item;
    try { clickable.click(); return true; } catch { return false; }
  };

  // 2) TEAM + CHANNEL path
  if (needleTeam && needleChannel) {
    // Find all team treeitems.
    const teamItems = [...document.querySelectorAll('[role="treeitem"][data-item-type="team"]')];
    let bestTeam = null;
    let bestTeamScore = 0;
    const teamCandidates = [];
    for (const ti of teamItems) {
      // Team label is in a span with id starting "title-team-list-item-"
      const titleEl = ti.querySelector('[id^="title-team-list-item-"]')
        || ti.querySelector('[class*="team-type-name"]')
        || ti.querySelector('span');
      const txt = titleEl?.textContent || '';
      const s = score(txt, needleTeam);
      teamCandidates.push({ text: txt.trim().slice(0, 80), score: s });
      if (s > bestTeamScore) { bestTeamScore = s; bestTeam = ti; }
    }
    if (!bestTeam || bestTeamScore < 25) {
      return {
        found: false, clicked: false, mode: 'team-channel',
        reason: 'team-not-found',
        expandedFolders: expanded,
        teamCandidates: teamCandidates.sort((a, b) => b.score - a.score).slice(0, 8),
      };
    }
    // Expand the team if collapsed.
    if (bestTeam.getAttribute('aria-expanded') === 'false') {
      const header = bestTeam.querySelector('[data-testid^="list-item-teams-and-channels"]') || bestTeam;
      try { header.click(); expanded++; } catch {}
    }

    // Match channels to this team by thread ID extracted from the team's
    // data-fui-tree-item-value (e.g. "OneGQL_Team|19:<id>@thread.tacv2").
    // Channel items have their team's thread ID embedded in their own
    // data-fui-tree-item-value and data-testid ("sc-channel-list-item-<id>"),
    // which is far more reliable than walking DOM siblings.
    const teamValue = bestTeam.getAttribute('data-fui-tree-item-value') || '';
    const teamIdMatch = teamValue.match(/19:[0-9a-f]+@thread\.(?:tacv2|skype)/i);
    const teamThreadId = teamIdMatch ? teamIdMatch[0] : null;

    let channelItems = [];
    if (teamThreadId) {
      channelItems = [...document.querySelectorAll('[role="treeitem"][data-item-type="channel"]')]
        .filter((ci) => {
          const v = ci.getAttribute('data-fui-tree-item-value') || '';
          const t = ci.getAttribute('data-testid') || '';
          return v.includes(teamThreadId) || t.includes(teamThreadId);
        });
    }
    // Fallback: positional (legacy behavior) if thread-ID pairing found nothing.
    if (channelItems.length === 0) {
      let group = null;
      let sib = bestTeam.nextElementSibling;
      while (sib && !group) {
        if (sib.matches?.('[role="group"]')) group = sib;
        sib = sib.nextElementSibling;
      }
      channelItems = group
        ? [...group.querySelectorAll('[role="treeitem"][data-item-type="channel"]')]
        : [];
    }
    let bestCh = null;
    let bestChScore = 0;
    const chCandidates = [];
    for (const ci of channelItems) {
      const titleEl = ci.querySelector('[id^="title-channel-list-item-"]')
        || ci.querySelector('[class*="channel-type-name"]')
        || ci.querySelector('span');
      const txt = titleEl?.textContent || '';
      const s = score(txt, needleChannel);
      chCandidates.push({ text: txt.trim().slice(0, 80), score: s });
      if (s > bestChScore) { bestChScore = s; bestCh = ci; }
    }
    if (!bestCh || bestChScore < 40) {
      // "See all channels" may hide it — try clicking that and re-query once.
      const seeAll = group?.querySelector('[data-testid$="-see-all-channel"], [data-testid*="seeall"]');
      if (seeAll) {
        try { seeAll.click(); } catch {}
      }
      return {
        found: false, clicked: false, mode: 'team-channel',
        reason: 'channel-not-found-in-team',
        teamMatch: (bestTeam.textContent || '').trim().slice(0, 80),
        expandedFolders: expanded,
        channelCandidates: chCandidates.sort((a, b) => b.score - a.score).slice(0, 10),
      };
    }
    const ok = clickTreeitem(bestCh);
    return {
      found: true, clicked: ok, mode: 'team-channel',
      teamMatch: (bestTeam.querySelector('[id^="title-team-list-item-"]')?.textContent || '').trim(),
      channelMatch: (bestCh.querySelector('[id^="title-channel-list-item-"]')?.textContent || '').trim(),
      teamScore: bestTeamScore, channelScore: bestChScore,
      expandedFolders: expanded,
    };
  }

  // 3) NAME-only path — search across all visible titles (chats + teams + channels).
  const needle = needleName;
  if (!needle) return { found: false, clicked: false, reason: 'empty-name' };

  const titleEls = document.querySelectorAll(
    '[data-testid="simple-collab-dnd-rail"] [id^="title-chat-list-item_"], '
    + '[data-testid="simple-collab-dnd-rail"] [id^="title-channel-list-item-"], '
    + '[data-testid="simple-collab-dnd-rail"] [id^="title-team-list-item-"], '
    + '[data-testid="simple-collab-dnd-rail"] [role="treeitem"] span[id^="title-"]'
  );
  const candidates = [];
  let best = null;
  let bestScore = 0;
  for (const el of titleEls) {
    const txt = el.textContent || '';
    const s = score(txt, needle);
    // Prefer chats and channels over teams for name-only queries.
    const ti = el.closest('[role="treeitem"]');
    const kind = ti?.getAttribute('data-item-type') || '';
    const weighted = s + (kind === 'chat' ? 3 : kind === 'channel' ? 1 : 0);
    candidates.push({ text: txt.trim().slice(0, 80), score: s, kind });
    if (weighted > bestScore) { bestScore = weighted; best = el; }
  }
  if (!best || bestScore < 25) {
    return {
      found: false, clicked: false, mode: 'name',
      expandedFolders: expanded,
      candidateCount: candidates.length,
      topCandidates: candidates.sort((a, b) => b.score - a.score).slice(0, 8),
    };
  }
  const ok = clickTreeitem(best);
  return {
    found: true, clicked: ok, mode: 'name',
    matchText: best.textContent?.trim() || '',
    score: bestScore,
    expandedFolders: expanded,
  };
}

// Ensure the Chat tab is active so the left rail is rendered. Presses Escape
// first to dismiss any open search overlay, then clicks the Chat app-bar item
// if the rail is currently missing.
export function ensureChatTabActive() {
  // Dismiss any modal/overlay the app may have opened.
  try {
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  } catch {}
  const railPresent = !!document.querySelector('[data-testid="simple-collab-dnd-rail"]');
  if (railPresent) return { clicked: false, railPresent: true };
  const selectors = [
    'button[aria-label="Chat" i]',
    '[data-tid="app-bar-chat"]',
    '[data-tid="app-bar-2a84919f-59d8-4441-a975-2a8c2643b741"]', // internal Chat app id
    'a[href*="/chat"]',
  ];
  for (const sel of selectors) {
    const el = document.querySelector(sel);
    if (el) {
      try { el.click(); return { clicked: true, via: sel, railPresent: false }; } catch {}
    }
  }
  return { clicked: false, railPresent: false, reason: 'chat-tab-not-found' };
}

// Scroll the sidebar chat/teams tree to load more virtualized items. Returns
// whether the scroll moved. Useful when a target isn't in the initial view.
export function scrollSidebarTree(delta) {
  const rail = document.querySelector('[data-testid="simple-collab-dnd-rail"]');
  if (!rail) return { ok: false, reason: 'no-rail' };
  // Walk up/down to find the scrollable ancestor or descendant.
  let el = rail;
  const find = (node) => {
    if (!node) return null;
    if (node.scrollHeight > node.clientHeight + 10) {
      const cs = getComputedStyle(node);
      if (cs.overflowY === 'auto' || cs.overflowY === 'scroll') return node;
    }
    return null;
  };
  let scroller = find(rail);
  if (!scroller) {
    let p = rail.parentElement;
    while (p && !scroller) { scroller = find(p); p = p.parentElement; }
  }
  if (!scroller) {
    for (const n of rail.querySelectorAll('*')) {
      scroller = find(n);
      if (scroller) break;
    }
  }
  if (!scroller) return { ok: false, reason: 'no-scroller' };
  const before = scroller.scrollTop;
  scroller.scrollTop = Math.min(scroller.scrollHeight, scroller.scrollTop + Number(delta || 600));
  return {
    ok: true, before, after: scroller.scrollTop,
    atBottom: scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 2,
    scrollHeight: scroller.scrollHeight,
  };
}



// DOM helpers for enterprise Microsoft Teams (teams.microsoft.com/v2).
//
// Each exported function is fully self-contained: selector banks + helpers are
// inlined inside each function body. Reason: Playwright's `page.evaluate(fn)`
// serializes only that one function's source, so module-scope constants and
// helper functions are NOT available in the browser context.

export function readHeader() {
  const hEl = document.querySelector(
    '[data-tid="chat-header-title"], [data-tid="chat-title"], [data-tid="channel-header-title"], [role="banner"] h1, [role="main"] h1, [role="main"] h2'
  );
  return { header: hEl?.textContent?.trim() || null, title: document.title };
}

export function collectPaneMessages() {
  const CHAT_CONTAINER_SELECTORS = [
    '.fui-Chat',
    '[class*="fui-Chat"]',
    '[data-tid="message-pane-list-surface"]',
    '.ts-message-list-container',
    '[data-tid="chat-pane"]',
    '[role="main"] [data-is-scrollable="true"]',
    '.message-list',
    '[class*="message-list"]',
    '[class*="MessageList"]',
    '[data-tid="messageListContainer"]',
    '[role="log"]',
  ];
  const MESSAGE_SELECTORS = [
    '[data-tid="chat-pane-item"]',
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
          index: idx, isSystemMessage: true, sender: '[System]',
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
      index: idx, isSystemMessage: false,
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
    '.ts-message-list-container', '[data-tid="chat-pane"]',
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
    const chat = getChatContainer();
    if (chat) {
      if (chat.scrollHeight > chat.clientHeight + 10) return chat;
      let p = chat.parentElement;
      while (p) {
        const cs = getComputedStyle(p);
        if ((cs.overflowY === 'auto' || cs.overflowY === 'scroll') && p.scrollHeight > p.clientHeight + 10) return p;
        p = p.parentElement;
      }
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
    '[data-tid="chat-pane"]', '[role="main"] [data-is-scrollable="true"]',
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

export function inspectPane() {
  const CHAT_CONTAINER_SELECTORS = [
    '.fui-Chat', '[class*="fui-Chat"]', '[data-tid="message-pane-list-surface"]',
    '.ts-message-list-container', '[data-tid="chat-pane"]',
    '[role="main"] [data-is-scrollable="true"]', '.message-list',
    '[class*="message-list"]', '[class*="MessageList"]',
    '[data-tid="messageListContainer"]', '[role="log"]',
  ];
  const MESSAGE_SELECTORS = [
    '[data-tid="chat-pane-item"]', '[data-testid="message-wrapper"]',
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

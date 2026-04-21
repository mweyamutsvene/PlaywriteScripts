// DOM helpers evaluated inside the Outlook page via page.evaluate.
// These functions are serialized by Playwright, so they must be self-contained
// (no closures over Node-side variables).

/**
 * Snapshot of every `[role="option"]` currently rendered in the virtualized
 * message list. Outlook lazily mounts/unmounts rows as the user scrolls, so
 * we call this repeatedly and merge in Node-land.
 */
export function collectVisibleItems() {
  const scroller =
    document.querySelector('[aria-label="Message list"] .customScrollBar[data-is-scrollable="true"]') ||
    [...document.querySelectorAll('[aria-label="Message list"] .customScrollBar')].find(s => s.querySelector('[role="option"]'));
  const listbox = document.querySelector('[aria-label="Message list"][role="listbox"], [aria-label="Message list"] [role="listbox"]');

  const dateTitleRe = /^\w{2,4}\s+\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s*(AM|PM)/i;
  const items = [];
  const options = document.querySelectorAll('[aria-label="Message list"] [role="option"][data-convid]');
  let totalSize = null;
  for (const opt of options) {
    if (totalSize == null) {
      const s = opt.getAttribute('aria-setsize');
      if (s) totalSize = Number(s);
    }
    let dateTitle = null;
    for (const el of opt.querySelectorAll('[title]')) {
      const t = el.getAttribute('title') || '';
      if (dateTitleRe.test(t)) { dateTitle = t; break; }
    }
    let senderName = null, senderEmail = null;
    for (const el of opt.querySelectorAll('[title]')) {
      const t = el.getAttribute('title') || '';
      if (t.includes('@') && !t.includes(' ')) {
        senderEmail = t;
        senderName = el.textContent?.trim() || null;
        break;
      }
    }
    const posIn = Number(opt.getAttribute('aria-posinset')) || null;
    items.push({
      id: opt.id,
      convId: opt.getAttribute('data-convid'),
      ariaLabel: opt.getAttribute('aria-label') || '',
      isUnread: (opt.getAttribute('aria-label') || '').startsWith('Unread'),
      senderName,
      senderEmail,
      dateTitle,
      posInSet: posIn,
    });
  }

  return {
    items,
    totalSize,
    scroll: scroller ? {
      scrollTop: scroller.scrollTop,
      scrollHeight: scroller.scrollHeight,
      clientHeight: scroller.clientHeight,
    } : null,
    hasScroller: !!scroller,
    hasListbox: !!listbox,
  };
}

export function scrollBy(delta) {
  const scroller =
    document.querySelector('[aria-label="Message list"] .customScrollBar[data-is-scrollable="true"]') ||
    [...document.querySelectorAll('[aria-label="Message list"] .customScrollBar')].find(s => s.querySelector('[role="option"]'));
  if (!scroller) return { ok: false };
  const before = scroller.scrollTop;
  scroller.scrollTop = Math.min(scroller.scrollTop + delta, scroller.scrollHeight);
  return {
    ok: true,
    before,
    after: scroller.scrollTop,
    scrollHeight: scroller.scrollHeight,
    clientHeight: scroller.clientHeight,
    atBottom: (scroller.scrollTop + scroller.clientHeight) >= (scroller.scrollHeight - 2),
  };
}

export function scrollToTop() {
  const scroller =
    document.querySelector('[aria-label="Message list"] .customScrollBar[data-is-scrollable="true"]') ||
    [...document.querySelectorAll('[aria-label="Message list"] .customScrollBar')].find(s => s.querySelector('[role="option"]'));
  if (scroller) scroller.scrollTop = 0;
  return !!scroller;
}

/** Reads the currently-open message in the Reading Pane. */
export function readOpenMessage() {
  const rp = document.querySelector('[aria-label="Reading Pane"]');
  if (!rp) return null;

  const subjectEl = rp.querySelector('[id$="_SUBJECT"] [title], [role="heading"][aria-level="3"]');
  const subject = subjectEl?.getAttribute('title') || subjectEl?.textContent?.trim() || null;

  const fromBtn = rp.querySelector('button[aria-label^="From:"]');
  const fromText = fromBtn?.textContent?.trim() || null;
  let fromName = null, fromEmail = null;
  if (fromText) {
    const m = fromText.match(/^(.*?)<([^>]+)>$/);
    if (m) { fromName = m[1].trim(); fromEmail = m[2].trim(); }
    else { fromName = fromText; }
  }

  const to = [...rp.querySelectorAll('[aria-label^="To:"] button')].map(b => b.textContent?.trim()).filter(Boolean);
  const cc = [...rp.querySelectorAll('[aria-label^="Cc:"] button')].map(b => b.textContent?.trim()).filter(Boolean);

  let dateTitle = null;
  for (const h of rp.querySelectorAll('[role="heading"]')) {
    const t = (h.textContent || '').trim();
    if (/^\w{2,4}\s+\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s*(AM|PM)/i.test(t)) {
      dateTitle = t; break;
    }
  }

  const bodyEl = rp.querySelector('[aria-label="Message body"]');
  const bodyText = bodyEl ? bodyEl.innerText.replace(/\u00a0/g, ' ').trim() : null;

  const convSubject = rp.querySelector('[id^="CONV_"][id$="_SUBJECT"]');
  const convKey = convSubject?.id || null;

  return { subject, fromName, fromEmail, to, cc, dateTitle, bodyText, convKey };
}

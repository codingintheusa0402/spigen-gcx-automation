// ==UserScript==
// @name         GChat Reply Suggest
// @namespace    https://spigen.com/gcx
// @version      3.7.0
// @description  Alt+G offers T3 Esc (deterministic ticket-forward, no AI) / Gratitude / Reminder templates in every Google Chat room by default; only in designated rooms does it suggest AI-generated reply sentences instead
// @author       Spigen GCX
// @updateURL    https://raw.githubusercontent.com/codingintheusa0402/spigen-gcx-automation/main/Browser_Extensions/tampermonkey_scripts/GChat%20Reply%20Suggest.user.js
// @downloadURL  https://raw.githubusercontent.com/codingintheusa0402/spigen-gcx-automation/main/Browser_Extensions/tampermonkey_scripts/GChat%20Reply%20Suggest.user.js
// @match        https://chat.google.com/*
// @match        https://spigenhelp.zendesk.com/agent/tickets/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @grant        GM_setValue
// @grant        GM_getValue
// @connect      localhost
// @connect      127.0.0.1
// @run-at       document-idle
// ==/UserScript==

(function () {
  "use strict";

  // As of 2026-08-04: the ticket-forward picker is the DEFAULT for every
  // room. AI-suggested replies are now the exception, only in rooms listed
  // here — keyed by Chat "space" ID, the /app/chat/<ID> segment of the URL.
  // This is NOT the room's display name: document.title is unreliable
  // (Chat rewrites it for unread counts, "X messaged you", etc.), so the
  // space ID is the only stable identifier. To add a room here (or to a
  // room that doesn't exist yet, like "GCX T2 ESC. Ticket 보고" below): once
  // it exists, open it, press Option+G once, then check the browser
  // console — this script logs the current space ID and name every time,
  // ready to copy in.
  const AI_SUGGEST_ROOM_IDS = {
    // "<space id>": "GCX T2 ESC. Ticket 보고", // TODO: fill in once this room is created
  };

  // Same custom-field IDs GCX Reply.user.js already uses (ZD constant there),
  // kept in sync so both scripts read the same ticket data the same way.
  const ZD = {
    ASIN: 360021934312,
    COUNTRY: 4513936822297,
    DEVICE: 360022185671,
    PRODUCT_NAME: 360022185891,
    INQUIRY_1: 360022182831, // "1차 Defect Reason or Inquiries"
    ESC_REASON: 900007557523, // "ESC. 사유"
    ORDER_ID: 360021934132,
  };
  // "Requested Channel type*" has no ID recorded anywhere in this repo, so
  // it's resolved by title from ticket_fields.json (see getFieldOptsCached).
  const CHANNEL_FIELD_TITLE_RE = /^Requested Channel type/i;
  const ORDER_ID_CHANNELS = ["Shopify EU"];

  const RECENT_TICKETS_KEY = "grs_recent_tickets";
  const MAX_RECENT_TICKETS = 10;
  const FIELD_OPTS_CACHE_KEY = "grs_zd_field_opts_cache_v2"; // v2: + channel field
  const FIELD_OPTS_TTL_MS = 24 * 60 * 60 * 1000;
  const PENDING_CONFIRM_PREFIX = "grs_pending_confirm_";
  const DAILY_INDEX_PREFIX = "grs_idx_";

  function todayKST() {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Seoul",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(new Date());
  }

  // =====================================================================
  // ZENDESK SIDE — records recently-visited tickets for the Chat side to
  // pick from. Runs only on spigenhelp.zendesk.com/agent/tickets/*.
  // =====================================================================
  async function getFieldOptsCached() {
    const cached = GM_getValue(FIELD_OPTS_CACHE_KEY, null);
    if (cached && Date.now() - cached.ts < FIELD_OPTS_TTL_MS) return cached.data;

    const res = await fetch("/api/v2/ticket_fields.json?locale=ko", { credentials: "include" });
    const data = await res.json();
    const wantedIds = [ZD.COUNTRY, ZD.DEVICE, ZD.PRODUCT_NAME, ZD.INQUIRY_1, ZD.ESC_REASON];
    const out = {};
    for (const f of data.ticket_fields) {
      const isChannel = CHANNEL_FIELD_TITLE_RE.test(f.title || "");
      if (isChannel) out.channelFieldId = f.id;
      if (!wantedIds.includes(f.id) && !isChannel) continue;
      const map = {};
      for (const o of f.custom_field_options || []) map[o.value] = o.name;
      out[f.id] = map;
    }
    GM_setValue(FIELD_OPTS_CACHE_KEY, { ts: Date.now(), data: out });
    return out;
  }

  async function recordCurrentTicket() {
    const m = location.pathname.match(/\/tickets\/(\d+)/);
    if (!m) return;
    const ticketId = m[1];

    try {
      const [ticketRes, fieldOpts] = await Promise.all([
        fetch(`/api/v2/tickets/${ticketId}.json`, { credentials: "include" }).then((r) => r.json()),
        getFieldOptsCached(),
      ]);
      const ticket = ticketRes.ticket || {};
      const cfMap = {};
      for (const f of ticket.custom_fields || []) cfMap[f.id] = f.value;

      const resolve = (fieldId) => {
        const val = cfMap[fieldId];
        if (!val) return "";
        const opts = fieldOpts[fieldId] || {};
        return opts[val] || val;
      };

      const entry = {
        id: ticketId,
        url: `https://spigenhelp.zendesk.com/agent/tickets/${ticketId}`,
        subject: ticket.subject || "(no subject)",
        country: resolve(ZD.COUNTRY),
        device: resolve(ZD.DEVICE),
        productName: resolve(ZD.PRODUCT_NAME),
        asin: cfMap[ZD.ASIN] || "",
        inquiryReason: resolve(ZD.INQUIRY_1) || resolve(ZD.ESC_REASON) || "",
        channel: fieldOpts.channelFieldId ? resolve(fieldOpts.channelFieldId) : "",
        orderId: String(cfMap[ZD.ORDER_ID] || "").trim(),
        visitedAt: Date.now(),
      };

      let list = GM_getValue(RECENT_TICKETS_KEY, []);
      list = list.filter((t) => t.id !== entry.id);
      list.unshift(entry);
      list = list.slice(0, MAX_RECENT_TICKETS);
      GM_setValue(RECENT_TICKETS_KEY, list);
      console.log(`[GChat Reply Suggest] recorded ticket #${ticketId}`, entry);
    } catch (e) {
      console.error(`[GChat Reply Suggest] failed to record ticket #${ticketId}`, e);
    }
  }

  // =====================================================================
  // CHAT SIDE
  // =====================================================================
  const BACKEND_URL = "http://127.0.0.1:8765/suggest";
  const TRANSCRIPT_CHAR_LIMIT = 2000;
  const BAR_ID = "gchat-reply-suggest-bar";

  GM_addStyle(`
    #${BAR_ID} {
      position: fixed;
      z-index: 999999;
      display: flex;
      flex-direction: column;
      gap: 6px;
      padding: 6px;
      max-width: 520px;
      pointer-events: none;
    }
    #${BAR_ID} .grs-row {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      pointer-events: none;
    }
    #${BAR_ID} .grs-chip {
      pointer-events: auto;
      background: #1a73e8;
      color: #fff;
      border-radius: 14px;
      padding: 6px 12px;
      font-size: 13px;
      font-family: "Google Sans", Roboto, Arial, sans-serif;
      cursor: pointer;
      box-shadow: 0 1px 4px rgba(0,0,0,0.3);
      max-width: 300px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    #${BAR_ID} .grs-chip:hover { background: #1558b3; }
    #${BAR_ID} .grs-chip.grs-loading { background: #5f6368; cursor: default; }
    #${BAR_ID} .grs-chip.grs-error { background: #d93025; cursor: default; white-space: normal; }
    #${BAR_ID} .grs-panel {
      pointer-events: auto;
      background: #fff;
      color: #202124;
      border-radius: 10px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.35);
      padding: 6px;
      max-width: 480px;
      max-height: 320px;
      overflow-y: auto;
      font-family: "Google Sans", Roboto, Arial, sans-serif;
      font-size: 13px;
    }
    #${BAR_ID} .grs-panel-title {
      font-weight: 600;
      padding: 4px 8px;
      color: #5f6368;
      font-size: 11px;
      text-transform: uppercase;
    }
    #${BAR_ID} .grs-item {
      padding: 8px;
      border-radius: 6px;
      cursor: pointer;
    }
    #${BAR_ID} .grs-item:hover { background: #f1f3f4; }
    #${BAR_ID} .grs-item.grs-item-default { background: #e8f0fe; }
    #${BAR_ID} .grs-item-subject { font-weight: 500; }
    #${BAR_ID} .grs-item-meta { color: #5f6368; font-size: 11px; margin-top: 2px; }
  `);

  function getComposeBox() {
    const boxes = Array.from(
      document.querySelectorAll('[role="textbox"][contenteditable="true"]')
    ).filter((el) => el.offsetParent !== null);
    return boxes.length ? boxes[boxes.length - 1] : null;
  }

  function getSpaceId() {
    const m = location.pathname.match(/\/chat\/([^/?#]+)/);
    return m ? m[1] : null;
  }

  // Best-effort display name, only used for logging/labels — never for the
  // AI-suggest-vs-ticket-forward decision itself (see AI_SUGGEST_ROOM_IDS
  // comment above).
  function getRoomName() {
    return (document.title || "")
      .replace(/^\(\d+\)\s*/, "")
      .replace(/\s*-\s*Chat\s*$/, "")
      .trim();
  }

  function isTemplateRoom() {
    return !(getSpaceId() in AI_SUGGEST_ROOM_IDS);
  }

  function getTranscriptTail() {
    const main = document.querySelector('[role="main"]');
    if (!main) return "";
    const text = (main.innerText || "").trim();
    return text.slice(-TRANSCRIPT_CHAR_LIMIT);
  }

  // Best-effort sender-name scrape. Google Chat's classes are hashed/rotate
  // on redeploy — if these stop matching, this degrades to an empty list
  // (manual @mention typing still works, nothing crashes).
  function getRecentSenders(limit = 8) {
    const main = document.querySelector('[role="main"]');
    if (!main) return [];
    const nodes = main.querySelectorAll(".MsqITd, .njhDLd.O5OMdc");
    const seen = [];
    for (let i = nodes.length - 1; i >= 0 && seen.length < limit; i--) {
      const name = (nodes[i].textContent || "").trim();
      if (name && !seen.includes(name)) seen.push(name);
    }
    return seen;
  }

  function removeBar() {
    pickerState = null;
    const existing = document.getElementById(BAR_ID);
    if (existing) existing.remove();
  }

  function positionBarAbove(el, bar) {
    const rect = el.getBoundingClientRect();
    bar.style.left = `${Math.max(8, rect.left)}px`;
    bar.style.bottom = `${Math.max(8, window.innerHeight - rect.top + 8)}px`;
  }

  function insertIntoCompose(box, text) {
    box.focus();
    const sel = window.getSelection();
    sel.removeAllRanges();
    const range = document.createRange();
    range.selectNodeContents(box);
    range.collapse(false); // cursor at end
    sel.addRange(range);
    document.execCommand("insertText", false, text);
  }

  // Chat's compose box enforces a Trusted Types CSP, so execCommand("insertHTML", ...)
  // throws. Instead: find the plain text we just inserted and wrap it with real
  // DOM nodes (createElement + appendChild — no HTML-string parsing involved,
  // so Trusted Types doesn't apply), then fire an input event so Chat's own
  // compose-state sync picks up the change. Verified this survives an actual
  // send (not just the draft view).
  function wrapSubstringWithElement(box, needle, makeWrapper) {
    if (!needle) return false;
    const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
      const idx = node.textContent.indexOf(needle);
      if (idx === -1) continue;
      const after = node.splitText(idx);
      after.splitText(needle.length);
      const wrapper = makeWrapper();
      after.parentNode.insertBefore(wrapper, after);
      wrapper.appendChild(after);
      return true;
    }
    return false;
  }

  function applyReferenceLineFormatting(box, refLineText, url) {
    wrapSubstringWithElement(box, refLineText, () => document.createElement("b"));
    wrapSubstringWithElement(box, url, () => {
      const a = document.createElement("a");
      a.href = url;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      return a;
    });
    box.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: true, inputType: "insertText" }));
  }

  function renderChips(box, suggestions) {
    removeBar();
    const bar = document.createElement("div");
    bar.id = BAR_ID;
    const row = document.createElement("div");
    row.className = "grs-row";
    suggestions.forEach((s) => {
      const chip = document.createElement("div");
      chip.className = "grs-chip";
      chip.textContent = s;
      chip.title = s;
      chip.addEventListener("click", () => {
        insertIntoCompose(box, s);
        removeBar();
      });
      row.appendChild(chip);
    });
    bar.appendChild(row);
    document.body.appendChild(bar);
    positionBarAbove(box, bar);
  }

  function renderStatus(box, text, cls) {
    removeBar();
    const bar = document.createElement("div");
    bar.id = BAR_ID;
    const row = document.createElement("div");
    row.className = "grs-row";
    const chip = document.createElement("div");
    chip.className = `grs-chip ${cls}`;
    chip.textContent = text;
    row.appendChild(chip);
    bar.appendChild(row);
    document.body.appendChild(bar);
    positionBarAbove(box, bar);
  }

  // ---- AI-suggest flow (non-template rooms) ----
  function requestAiSuggestions() {
    const box = getComposeBox();
    if (!box) return;

    const transcript = getTranscriptTail();
    if (!transcript) {
      renderStatus(box, "No conversation found to suggest from", "grs-error");
      setTimeout(removeBar, 2500);
      return;
    }

    renderStatus(box, "Thinking…", "grs-loading");

    GM_xmlhttpRequest({
      method: "POST",
      url: BACKEND_URL,
      headers: { "Content-Type": "application/json" },
      data: JSON.stringify({ transcript }),
      timeout: 30000,
      onload: (res) => {
        try {
          const body = JSON.parse(res.responseText);
          if (body.suggestions && body.suggestions.length) {
            renderChips(box, body.suggestions);
          } else {
            renderStatus(box, body.error || "No suggestions returned", "grs-error");
            setTimeout(removeBar, 3000);
          }
        } catch (e) {
          renderStatus(box, "Bad response from backend", "grs-error");
          setTimeout(removeBar, 3000);
        }
      },
      onerror: () => {
        renderStatus(box, "Backend not running (localhost:8765)", "grs-error");
        setTimeout(removeBar, 3500);
      },
      ontimeout: () => {
        renderStatus(box, "Backend timed out", "grs-error");
        setTimeout(removeBar, 3000);
      },
    });
  }

  // ---- Ticket-forward template flow (template rooms) ----
  function referenceLineText(t, index) {
    // Omit the bracket entirely when there's no ASIN on file, rather than
    // producing a malformed/empty "[]" — this also matters for
    // getThreadRootInfo()'s detection below, which no longer depends on the
    // bracket being present.
    const asinPart = t.asin ? ` [${t.asin}]` : "";
    return `${index} ${t.country} ${t.device}용 ${t.productName}${asinPart}`;
  }

  function referenceBlock(t, index) {
    return `${referenceLineText(t, index)}\n${t.url}`;
  }

  // No mentionName parameter here — a real, notification-triggering @mention
  // can only come from Chat's own People autocomplete, which requires a
  // genuinely user-trusted keydown — verified two ways: (1) a synthetic JS
  // keydown/beforeinput/keyup "@" sequence produced zero DOM changes
  // anywhere on the page, and (2) even a genuinely trusted CDP-level
  // keystroke (the mechanism Playwright/Puppeteer use, isTrusted: true,
  // indistinguishable from real hardware at the browser's own trust layer)
  // still didn't trigger Chat's popup — so this isn't just an isTrusted
  // check, it's a hard wall with no scriptable path around it. Given that,
  // mentionName is embedded as plain text: it won't actually notify the
  // person, but it's the least-friction option and was chosen deliberately
  // over the alternative (place cursor + require the user to type the real
  // "@" themselves) after weighing both.
  const HONORIFIC_TITLES = ["리더", "파트장", "프로"];
  const DEFAULT_HONORIFIC = "프로";

  // Zendesk's own option labels are often "{English category}_{Korean detail}"
  // (e.g. "(Sales Inquiry)_재입고관련문의", "(Urgent)_리스팅오류") — only the
  // part after the last "_" is wanted in the template; the English/bracketed
  // category prefix gets dropped. Falls back to the whole string if there's
  // no underscore, and to "문의" if there's no label at all.
  function extractReasonLabel(raw) {
    const s = (raw || "").trim();
    if (!s) return "문의";
    const idx = s.lastIndexOf("_");
    const tail = (idx === -1 ? s : s.slice(idx + 1)).trim();
    return tail.replace(/^\((.*)\)$/, "$1") || "문의";
  }

  // Shopify EU tickets carry a Shopify order number (e.g. "1763" / "#1763")
  // in the Order ID field — append it so the receiver can look it up.
  function orderIdSuffix(t) {
    const norm = (v) => String(v || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    if (!ORDER_ID_CHANNELS.map(norm).includes(norm(t.channel))) return "";
    const id = (t.orderId || "").replace(/^#+/, "").trim();
    return id ? ` (주문번호: #${id})` : "";
  }

  function buildForwardText(t, index, mentionName, title) {
    const reason = extractReasonLabel(t.inquiryReason);
    const prefix = mentionName ? `@${mentionName} ` : "";
    const honorific = title || DEFAULT_HONORIFIC;
    return `${prefix}안녕하세요 ${honorific}님, 담당하시는 제품 관련 (${reason}) 문의가 들어와 전달드립니다. 확인 후 회신해 주시면 감사하겠습니다!${orderIdSuffix(t)}\n\n${referenceBlock(t, index)}`;
  }

  function buildConfirmText(pending, mentionName, title) {
    const prefix = mentionName ? `@${mentionName} ` : "";
    const honorific = title || DEFAULT_HONORIFIC;
    return `${prefix}확인 감사합니다 ${honorific}님. 주신 답변 확인 후 처리하도록 하겠습니다!\n\n${referenceBlock(pending, pending.index)}`;
  }

  // ---- Thread-context flows (Gratitude / Reminder) ----
  // These read the @mention + honorific from the T3 Esc root message the
  // currently-open thread is attached to, rather than asking again. Finding
  // "the root message" by DOM container/geometry turned out to be fragile
  // (the Thread side panel's exact ancestor boundary doesn't stand out
  // reliably), so instead this searches for our own script's distinctive
  // bold reference-line signature ("<n> <2-letter country> ... [ASIN]"),
  // restricted to the right half of the viewport (the thread panel sits to
  // the right of the main room panel — verified live: main-panel matches
  // sit at a fixed x, thread-panel matches at a much larger x), and takes
  // the topmost such match, since the root message is pinned above the
  // reply list. This only works for threads whose root was actually sent
  // via this script's T3 Esc flow (has the bold ref line) — verified this
  // finds and correctly parses real messages from actual usage.
  function getThreadRootInfo() {
    // Anchored on "<index> <2-letter country> ... 용 ..." — the "용"
    // particle (device-purpose suffix, e.g. "iPhone 15용") always sits
    // between device and product name in this template, so it identifies
    // our own reference line regardless of whether an ASIN bracket is
    // present. An earlier version required a trailing "[ASIN]" match here,
    // which silently failed to find threads on tickets with no ASIN on
    // file — referenceLineText() used to always append the bracket even
    // when empty, and this regex required 6-12 characters inside it.
    const refLineRe = /^\d+\s+[A-Z]{2}\s+.+용\s+.+/;
    const vw = window.innerWidth;
    const bolds = Array.from(document.querySelectorAll("b")).filter(
      (b) => b.offsetParent !== null && refLineRe.test((b.textContent || "").trim())
    );
    const rightSide = bolds.filter((b) => b.getBoundingClientRect().left > vw / 2);
    if (!rightSide.length) return null;
    const topmost = rightSide.sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top)[0];

    let node = topmost;
    for (let i = 0; i < 6 && node.parentElement; i++) {
      node = node.parentElement;
      if ((node.innerText || "").length > 40) break;
    }
    const text = node.innerText || "";

    const withMention = text.match(/@\s*([^\n@]+?)\s*\n?\s*안녕하세요\s+(리더|파트장|프로)님/);
    if (withMention) {
      return { name: withMention[1].replace(/\s+/g, " ").trim(), honorific: withMention[2] };
    }
    const noMention = text.match(/안녕하세요\s+(리더|파트장|프로)님/);
    return noMention ? { name: null, honorific: noMention[1] } : null;
  }

  function buildGratitudeText(info) {
    const namePart = info.name ? `${info.name} ` : "";
    return `${namePart}${info.honorific}님, 확인 감사합니다. 주신 메모 확인 후 처리하도록 하겠습니다!`;
  }

  function buildReminderText(info) {
    const namePart = info.name ? `${info.name} ` : "";
    return `안녕하세요 ${namePart}${info.honorific}님. 해당 티켓에 대한 답변이 아직 없어서 업무 마감차 리마인드챗드립니다. 시간 되실 때 확인 한 번 부탁드립니다. 감사합니다!`;
  }

  function insertThreadFollowup(kind) {
    const box = getComposeBox();
    if (!box) return;
    // If the thread's T3 Esc root can't be found (thread on an older/manual
    // message, root scrolled out, etc.), fall back to a name-less version
    // with the default honorific rather than an error chip — per explicit
    // request: "프로님, 확인 감사합니다. ..." is still a usable reply.
    const info = getThreadRootInfo() || { name: null, honorific: DEFAULT_HONORIFIC };
    const text = kind === "gratitude" ? buildGratitudeText(info) : buildReminderText(info);
    insertIntoCompose(box, text);
  }

  // Used by the "drafted" -> "sent" watcher transition (see
  // setupPendingConfirmWatcher) to check whether a forward has actually been
  // posted yet. The compose box lives INSIDE [role="main"] (verified:
  // main.contains(box) is true), so a naive main.innerText/Range scan picks
  // up whatever's currently drafted but not yet sent — not just
  // actually-posted messages. This excludes the compose box's subtree so
  // only real message content is considered. sinceNode, if given, restricts
  // the start boundary — otherwise scans from the very start of main.
  function getMessageOnlyText(main, sinceNode) {
    const composeBox = getComposeBox();
    const range = document.createRange();
    if (sinceNode) {
      range.setStartBefore(sinceNode);
    } else {
      range.selectNodeContents(main);
      range.collapse(true);
    }
    if (composeBox && main.contains(composeBox)) {
      range.setEndBefore(composeBox);
    } else {
      range.setEndAfter(main.lastChild || main);
    }
    return range.toString();
  }

  const MAX_DAILY_INDEX = 15;

  // Index is now a user pick, not an auto-scanned guess (the earlier DOM-scan
  // approach — searching the room's actual messages for the highest
  // reference-line index used — was replaced per explicit request: the user
  // picks from a 1..15 dropdown instead, defaulting to "next after whatever
  // was last confirmed sent today in this room"). This local counter is the
  // only source of truth now — keyed by room + day (todayKST()), so it
  // resets naturally at midnight KST with no separate reset logic needed.
  //
  // Read-only: what to default-select in the index dropdown. Does NOT
  // persist anything — safe to call on every draft/re-draft/re-open.
  // Persisting here would mean re-opening the picker and picking a
  // different ticket, or aborting and redrafting, silently burns index
  // numbers that were never actually sent.
  function peekNextIndexForToday(spaceId) {
    const key = DAILY_INDEX_PREFIX + spaceId + "_" + todayKST();
    const localMax = GM_getValue(key, 0);
    return Math.min(localMax + 1, MAX_DAILY_INDEX);
  }

  // Only call once a draft is CONFIRMED actually sent (see the "drafted" ->
  // "sent" transition in setupPendingConfirmWatcher) — this is the only
  // place the counter should advance, so a merely-drafted-then-abandoned
  // pick doesn't push the next default forward.
  function commitIndexUsed(spaceId, index) {
    const key = DAILY_INDEX_PREFIX + spaceId + "_" + todayKST();
    if (index > GM_getValue(key, 0)) GM_setValue(key, index);
  }

  // items = [1..15], default-highlighted on whatever peekNextIndexForToday()
  // suggests — e.g. if 1 was already sent today, 2 is pre-highlighted, but
  // 1 still sits right above it in the list (↑ once reaches it) in case you
  // need to reuse it (a redo, a correction, etc.).
  function showIndexPicker(defaultIndex, onPicked) {
    const box = getComposeBox();
    if (!box) return;
    const items = Array.from({ length: MAX_DAILY_INDEX }, (_, i) => i + 1);
    const clamped = Math.min(Math.max(defaultIndex, 1), MAX_DAILY_INDEX);
    openPicker(
      box,
      "Which index? (↑↓, Enter to pick)",
      items,
      clamped - 1,
      (el, n) => {
        el.textContent = String(n);
      },
      onPicked
    );
  }

  function relativeTime(ts) {
    const mins = Math.round((Date.now() - ts) / 60000);
    if (mins < 1) return "just now";
    if (mins < 60) return `${mins}m ago`;
    const hrs = Math.round(mins / 60);
    if (hrs < 24) return `${hrs}h ago`;
    return `${Math.round(hrs / 24)}d ago`;
  }

  // Active picker state (null when no picker is open). Generic across the
  // ticket picker and the member/@mention picker so both share one
  // arrow-key-nav + Enter/click/Escape implementation.
  // Shape: { items, selectedIndex, box, onConfirm(item) }
  let pickerState = null;

  function confirmPickerSelection() {
    if (!pickerState) return;
    const { items, selectedIndex } = pickerState;
    const item = items[selectedIndex];
    const onConfirm = pickerState.onConfirm;
    removeBar(); // clears pickerState
    onConfirm(item);
  }

  function highlightPickerSelection() {
    if (!pickerState) return;
    const els = document.querySelectorAll(`#${BAR_ID} .grs-item`);
    els.forEach((el, i) => {
      el.classList.toggle("grs-item-default", i === pickerState.selectedIndex);
    });
    const active = els[pickerState.selectedIndex];
    if (active) active.scrollIntoView({ block: "nearest" });
  }

  function openPicker(box, title, items, defaultIndex, renderItem, onConfirm, opts) {
    removeBar();
    const wrap = !!(opts && opts.wrap);
    pickerState = { items, selectedIndex: defaultIndex, box, onConfirm, wrap };

    const bar = document.createElement("div");
    bar.id = BAR_ID;
    const panel = document.createElement("div");
    panel.className = "grs-panel";

    const titleEl = document.createElement("div");
    titleEl.className = "grs-panel-title";
    titleEl.textContent = title;
    panel.appendChild(titleEl);

    items.forEach((item, i) => {
      const el = document.createElement("div");
      el.className = "grs-item" + (i === defaultIndex ? " grs-item-default" : "");
      renderItem(el, item, i);
      el.addEventListener("click", () => {
        if (pickerState) pickerState.selectedIndex = i;
        confirmPickerSelection();
      });
      panel.appendChild(el);
    });

    bar.appendChild(panel);
    document.body.appendChild(bar);
    positionBarAbove(box, bar);
  }

  const NO_MENTION = "(no mention)";

  function insertForwardWithMention(t, mentionName, honorific, index) {
    const box = getComposeBox();
    if (!box) return;
    const spaceId = getSpaceId();
    const refLineText = referenceLineText(t, index);
    insertIntoCompose(box, buildForwardText(t, index, mentionName, honorific));
    applyReferenceLineFormatting(box, refLineText, t.url);
    // stage: "drafted" — NOT yet awaiting a reply. The watcher below only
    // starts watching once this text actually shows up as a sent message
    // in the room (see setupPendingConfirmWatcher), otherwise an unrelated
    // message arriving while you're still editing the draft gets mistaken
    // for someone replying to a message you haven't sent yet.
    GM_setValue(PENDING_CONFIRM_PREFIX + spaceId, {
      index,
      country: t.country,
      device: t.device,
      productName: t.productName,
      asin: t.asin,
      url: t.url,
      refLineText,
      stage: "drafted",
      setAt: Date.now(),
    });
  }

  function showMentionPicker(title, defaultSender, onPicked) {
    const box = getComposeBox();
    if (!box) return;
    const senders = getRecentSenders();
    const names = defaultSender && !senders.includes(defaultSender) ? [defaultSender, ...senders] : senders;
    const items = [...names, NO_MENTION];
    openPicker(
      box,
      title,
      items,
      0,
      (el, name) => {
        el.textContent = name;
      },
      (name) => onPicked(name === NO_MENTION ? null : name),
      { wrap: true }
    );
  }

  function showHonorificPicker(onPicked) {
    const box = getComposeBox();
    if (!box) return;
    const defaultIndex = Math.max(0, HONORIFIC_TITLES.indexOf(DEFAULT_HONORIFIC));
    openPicker(
      box,
      "Which honorific? (↑↓, Enter to pick)",
      HONORIFIC_TITLES,
      defaultIndex,
      (el, honorific) => {
        el.textContent = `${honorific}님`;
      },
      onPicked
    );
  }

  function showTicketPicker() {
    const box = getComposeBox();
    if (!box) return;

    const tickets = GM_getValue(RECENT_TICKETS_KEY, []);
    if (!tickets.length) {
      renderStatus(box, "No recent tickets — open a Zendesk ticket first", "grs-error");
      setTimeout(removeBar, 3500);
      return;
    }

    openPicker(
      box,
      "Forward which ticket? (↑↓ to browse, Enter to pick)",
      tickets,
      0,
      (el, t) => {
        const subject = document.createElement("div");
        subject.className = "grs-item-subject";
        subject.textContent = `#${t.id} — ${t.subject}`;
        const meta = document.createElement("div");
        meta.className = "grs-item-meta";
        meta.textContent = `${t.country || "?"} · ${t.device || "?"} · ${t.productName || "?"} · ${t.asin || "?"} · ${relativeTime(t.visitedAt)}`;
        el.appendChild(subject);
        el.appendChild(meta);
      },
      (t) => {
        showMentionPicker("Mention who? (↑↓, Enter to pick)", null, (mentionName) => {
          showHonorificPicker((honorific) => {
            const spaceId = getSpaceId();
            showIndexPicker(peekNextIndexForToday(spaceId), (index) => {
              insertForwardWithMention(t, mentionName, honorific, index);
            });
          });
        });
      }
    );
  }

  function showConfirmSuggestion(pending, defaultSender) {
    showMentionPicker("Reply as confirm to whom? (↑↓, Enter to pick)", defaultSender, (mentionName) => {
      showHonorificPicker((honorific) => {
        const box = getComposeBox();
        if (!box) return;
        insertIntoCompose(box, buildConfirmText(pending, mentionName, honorific));
        applyReferenceLineFormatting(box, referenceLineText(pending, pending.index), pending.url);
        // Done watching this forward — clear it so the watcher stops here
        // instead of re-triggering on the next unrelated message too.
        GM_setValue(PENDING_CONFIRM_PREFIX + getSpaceId(), null);
      });
    });
  }

  // Watches for a new incoming message while a forward is pending
  // confirmation in a template room, and offers the confirm-reply chip.
  // Two-stage: "drafted" → "sent". Only messages that arrive AFTER the
  // forward is actually visible in the room count as a possible reply —
  // otherwise something unrelated arriving while you're still drafting (or
  // typing the real @mention) gets mistaken for a reply to a message you
  // haven't sent yet, and whatever you click gets appended onto your draft.
  function setupPendingConfirmWatcher() {
    const main = document.querySelector('[role="main"]');
    if (!main) return;

    let debounceTimer = null;
    const observer = new MutationObserver(() => {
      if (!isTemplateRoom()) return;
      const spaceId = getSpaceId();
      const pending = GM_getValue(PENDING_CONFIRM_PREFIX + spaceId, null);
      if (!pending) return;

      if (pending.stage === "drafted") {
        // getMessageOnlyText, not main.innerText — the compose box lives
        // inside main, so a plain scan would "detect" the draft itself as
        // already sent the instant it's typed.
        const text = getMessageOnlyText(main);
        if (pending.refLineText && text.includes(pending.refLineText)) {
          GM_setValue(PENDING_CONFIRM_PREFIX + spaceId, { ...pending, stage: "sent", sentAt: Date.now() });
          // Only now, confirmed actually sent, does this index count as used.
          commitIndexUsed(spaceId, pending.index);
        }
        return;
      }

      if (pending.stage !== "sent") return;
      // Give the DOM a moment to fully settle after our own message
      // renders — the tail end of that same render is still "mutations",
      // not a genuinely new/different message.
      if (Date.now() - (pending.sentAt || 0) < 1500) return;

      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        const senders = getRecentSenders(1);
        const latestSender = senders[0] || "";
        showConfirmSuggestion(pending, latestSender);
      }, 800);
    });
    observer.observe(main, { childList: true, subtree: true });
  }

  function showActionPicker() {
    const box = getComposeBox();
    if (!box) return;
    openPicker(
      box,
      "Which action? (↑↓, Enter to pick)",
      ["T3 Esc", "Gratitude", "Reminder"],
      0,
      (el, action) => {
        el.textContent = action;
      },
      (action) => {
        if (action === "T3 Esc") {
          showTicketPicker();
        } else {
          insertThreadFollowup(action === "Gratitude" ? "gratitude" : "reminder");
        }
      }
    );
  }

  function requestSuggestions() {
    // Self-service room-ID discovery: to make a room use AI-suggest instead
    // of the (now default) ticket-forward picker, open it, press Option+G
    // once, then check this log line and copy the id into
    // AI_SUGGEST_ROOM_IDS above.
    console.log(`[GChat Reply Suggest] space id: ${getSpaceId()} | name: ${getRoomName()} | isTemplateRoom: ${isTemplateRoom()}`);

    if (isTemplateRoom()) {
      showActionPicker();
    } else {
      requestAiSuggestions();
    }
  }

  function initChatSide() {
    // Capture phase (the `true` below) + stopImmediatePropagation: while a
    // picker is open, ArrowUp/ArrowDown/Enter/Escape must never reach Chat's
    // own keyboard handling. Verified this was a real bug, not theoretical:
    // preventDefault() alone only cancels the browser's default action for
    // the key — it does NOT stop other listeners (Chat's own) from also
    // reacting to the same event. Chat treats ArrowUp with an empty compose
    // box as "edit my last sent message" (a common chat-app convention), so
    // navigating the picker with ↑ was silently also flipping the previous
    // sent message into edit mode underneath it. Listening on the capture
    // phase means we see the event before Chat's own bubbling-phase
    // listeners do at all, regardless of registration order (Chat's script
    // loads long before this one, at document-idle, so on the bubbling
    // phase alone Chat's listener would win the race).
    document.addEventListener(
      "keydown",
      (e) => {
        if (pickerState) {
          if (e.key === "ArrowDown") {
            e.preventDefault();
            e.stopImmediatePropagation();
            const n = pickerState.items.length;
            pickerState.selectedIndex = pickerState.wrap
              ? (pickerState.selectedIndex + 1) % n
              : Math.min(pickerState.selectedIndex + 1, n - 1);
            highlightPickerSelection();
            return;
          }
          if (e.key === "ArrowUp") {
            e.preventDefault();
            e.stopImmediatePropagation();
            const n = pickerState.items.length;
            pickerState.selectedIndex = pickerState.wrap
              ? (pickerState.selectedIndex - 1 + n) % n
              : Math.max(pickerState.selectedIndex - 1, 0);
            highlightPickerSelection();
            return;
          }
          if (e.key === "Enter") {
            e.preventDefault();
            e.stopImmediatePropagation();
            confirmPickerSelection();
            return;
          }
          if (e.key === "Escape") {
            e.preventDefault();
            e.stopImmediatePropagation();
            removeBar();
            return;
          }
        }

        // Use e.code (physical key), not e.key: on Mac, Option+G produces "©"
        // as e.key (dead-key/diacritic composition), not the letter "g".
        if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && e.code === "KeyG") {
          e.preventDefault();
          e.stopImmediatePropagation();
          requestSuggestions();
        }
        if (e.key === "Escape") removeBar();
      },
      true
    );

    // [role="main"] renders after the conversation loads; retry briefly.
    let tries = 0;
    const tryObserve = setInterval(() => {
      tries++;
      if (document.querySelector('[role="main"]')) {
        setupPendingConfirmWatcher();
        clearInterval(tryObserve);
      } else if (tries > 40) {
        clearInterval(tryObserve);
      }
    }, 500);
  }

  // =====================================================================
  // ENTRY POINT
  // =====================================================================
  if (location.hostname === "chat.google.com") {
    initChatSide();
  } else if (location.hostname === "spigenhelp.zendesk.com") {
    // Zendesk is a single-page app — clicking between tickets (a list, a
    // "next ticket" button, etc.) changes the URL without a full page
    // reload, so a one-shot call here only ever records whichever ticket
    // happened to be open at the moment the script first loaded. Poll for
    // the ticket ID in the URL to change and re-record each time, instead.
    let lastTicketId = null;
    const checkForTicketChange = () => {
      const m = location.pathname.match(/\/tickets\/(\d+)/);
      const id = m ? m[1] : null;
      if (id && id !== lastTicketId) {
        lastTicketId = id;
        recordCurrentTicket();
      }
    };
    checkForTicketChange();
    setInterval(checkForTicketChange, 1500);
  }
})();

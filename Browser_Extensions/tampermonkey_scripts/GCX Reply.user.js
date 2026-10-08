// ==UserScript==
// @name         GCX Reply
// @namespace    https://spigen.com/gcx
// @version      3.8.0
// @description  Amazon order data via GAS web app + Spigen product info + Zendesk auto-fill
// @author       Spigen GCX
// @updateURL    https://raw.githubusercontent.com/codingintheusa0402/spigen-gcx-automation/main/Browser_Extensions/tampermonkey_scripts/GCX%20Reply.user.js
// @downloadURL  https://raw.githubusercontent.com/codingintheusa0402/spigen-gcx-automation/main/Browser_Extensions/tampermonkey_scripts/GCX%20Reply.user.js
// @match        https://spigenhelp.zendesk.com/agent/*
// @match        https://sellercentral.amazon.*/mcf/orders/create-order*
// @match        https://sellercentral.amazon.com/mcf/orders/create-order*
// @match        https://sellercentral.amazon.co.uk/mcf/orders/create-order*
// @match        https://sellercentral.amazon.de/mcf/orders/create-order*
// @match        https://sellercentral.amazon.fr/mcf/orders/create-order*
// @match        https://sellercentral.amazon.it/mcf/orders/create-order*
// @match        https://sellercentral.amazon.es/mcf/orders/create-order*
// @match        https://sellercentral-europe.amazon.*/mcf/orders/create-order*
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @grant        unsafeWindow
// @grant        GM_setValue
// @grant        GM_getValue
// @run-at       document-idle
// @connect      *
// html2canvas removed — Liquid Glass WebGL deprecated
// ==/UserScript==

(function () {
  'use strict';

  // Who's using this browser for MCF sends — written to "MCF 발송 로그"'s U col
  // (담당자) so the log reflects the actual agent instead of a hardcoded name.
  // Must be declared HERE, before the SC-page early-return right below — that
  // return skips everything after it, including a `const` further down the
  // file, so a deferred call into initMcfPage_'s autoFillFromHash_ (which
  // references these) hit a genuine ReferenceError (temporal dead zone) on
  // the SC page specifically, silently swallowed by that function's own
  // catch block. Confirmed live via injected diagnostic logging: "Cannot
  // access 'MCF_PERSON_OPTIONS' before initialization" — exact same class of
  // bug the comment on `_abmRelayHookInstalled` below already warns about,
  // just hit for a different variable on the other page branch. v3.6.9's fix
  // moved the prompt+storage to the SC page but declared these too late for
  // that same page to actually reach them.
  const MCF_PERSON_OPTIONS = ['양숙랑', '나아름', '임영신', '김지우', 'TCK'];
  function getMcfPerson_() { try { return GM_getValue('gcx_mcf_person', '') || ''; } catch (_) { return ''; } }
  function setMcfPerson_(name) { try { GM_setValue('gcx_mcf_person', name || ''); } catch (_) {} }

  // ── Seller Central MCF page: run embedded autofill instead of Zendesk panel ──
  if (location.hostname.includes('sellercentral') && location.pathname.includes('/mcf/orders/create-order')) {
    initMcfPage_();
    return;
  }

  // Install the ABM auto-relay fetch hook as early as possible on any Zendesk
  // page — independent of panel/init state, since a reply can be submitted
  // before the floating panel ever finishes mounting. Must NOT be gated on
  // the current pathname looking like a ticket URL: Zendesk's Agent
  // Workspace is a SPA, so if this tab's first real (full) page load landed
  // on the dashboard/views page and the agent then clicked into tickets via
  // client-side routing, a pathname gate here would never re-evaluate — the
  // hook would stay uninstalled for the tab's entire lifetime until a hard
  // reload happened to land directly on a ticket URL. The hook itself is
  // inert (only reacts to an actual UpdateTicketMutation POST), so there's
  // no cost to installing it unconditionally on every Zendesk page load.
  // `_abmRelayHookInstalled` must be declared here (not down near
  // installAbmRelayHook_'s other definition) — `let` has a temporal dead
  // zone, so calling the function this early while the variable is still
  // declared further down the file throws ReferenceError before anything
  // else in the script (panel, toggle button, etc.) ever runs.
  let _abmRelayHookInstalled = false;
  if (location.hostname.includes('zendesk.com')) {
    installAbmRelayHook_();
  }

  const GAS_URL    = 'https://script.google.com/macros/s/AKfycbw2Vdwk197LXB6oUAzuHS8sKamD5uqKZJDLvcHzbftWJk-M65XV1fAnTqiZo7ZEm4hk/exec';
  const SHEET_URL  = 'https://docs.google.com/spreadsheets/d/1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo/edit?gid=0#gid=0';
  const ORDER_RE   = /\b(\d{3}-\d{7}-\d{7})\b/g;
  const ASIN_RE    = /\b(B[A-Z0-9]{9})\b/g;
  const PANEL_ID   = 'sp-order-panel';
  const SHEET_COLS = ['SKU', '모델명', '브랜드', '제조사명', '기종명', '색상명', '대분류', '생산업체', '원산지정보'];

  // Any Zendesk tab doubles as an opportunistic worker for the shared
  // pending-ABM-relay queue — see sweepAbmRelayFailures_ for why this can't
  // be a real always-on backend job instead. Deferred via setTimeout so it
  // never competes with the panel's own init for the first render.
  if (location.hostname.includes('zendesk.com')) {
    if (typeof setTimeout === 'function') setTimeout(() => { sweepAbmRelayFailures_(); }, 3000);
    if (typeof setInterval === 'function') setInterval(() => { sweepAbmRelayFailures_(); }, 5 * 60 * 1000);
    // SC login-session check: both the ABM relay and the retry sweep need a
    // live cookie session on each marketplace's own SC domain — alert the
    // agent up front instead of letting sends fail silently for hours.
    if (typeof setTimeout === 'function') setTimeout(() => { checkScSessions_(); }, 4500);
    if (typeof setInterval === 'function') setInterval(() => { checkScSessions_(); }, 15 * 60 * 1000);
  }

  // ── Version update popup ────────────────────────────────────────────────
  // Keep CURRENT_VERSION in sync with the @version header above, and add a
  // CHANGELOG_ entry, every time the version is bumped. Empty/missing entries
  // are skipped silently (no popup shown).
  const CURRENT_VERSION = '3.8.0';
  // Stale per-version entries pruned — CHANGELOG_[CURRENT_VERSION] is the
  // only access pattern, so every past entry becomes dead weight the moment
  // its version is superseded. Add a fresh '<version>': [...] entry here on
  // the next version bump that needs a "what's new" popup.
  const CHANGELOG_ = {
    // Full per-version log since v3.0.0, per user request (not the earlier
    // capped-at-10 consolidated style) — kept as one string per version so
    // it's easy to keep appending to on future bumps.
    '3.2.1': [
      'v3.0.1 — MCF ASIN 선택 팝업에서 고른 ASIN이 무시되고 티켓 본문에서 처음 발견된 ASIN이 적용되던 버그 수정',
      'v3.0.2 — MCF 주소 자동입력 버그 수정: 주소와 무관한 번호 목록(예: 내부 처리사항)이 있을 경우 실제 주소 대신 해당 목록이 잘못 채워지던 문제 해결',
      'v3.0.3 — 업데이트 안내 팝업에 v3.0.1+v3.0.2 changelog 등록',
      'v3.0.4 — Confirm Auto-Fill 개선: Purchase Date를 yyyy-mm-dd로 표시, 체크 해제된 행 클릭 시 가독성 문제 수정, Customer Full Name DOM 폴백 추가, \'전체 주문\'/\'전체 환불\'을 주문 건수 대신 총 구매/환불 수량 기준으로 변경',
      'v3.0.5 — Purchase Date 입력 필드가 OS 로케일에 따라 다르게 표시되던 문제 수정 (항상 yyyy-mm-dd 표시 + 달력 선택 가능)',
      'v3.0.6 — Customer Full Name DOM 폴백 버그 수정: Customer context 패널이 지연 렌더링되는 문제라 자동으로 패널을 먼저 열도록 개선',
      'v3.0.7 — Auto-Fill 실행 후 Apps 탭이 아닌 Customer context 탭에 머물러 있던 버그 수정',
      'v3.1.0 — Seller Central Seller Notes를 GCX Reply 패널에서 자동으로 확인 가능하도록 추가',
      'v3.1.1 — Notes/Seller Notes 양방향 편집 기능 추가 (이후 3.1.3에서 다시 읽기 전용으로 롤백)',
      'v3.1.2 — Notes 양방향 동기화 버그 수정, Seller Notes가 계속 로딩 중으로 보이던 버그 수정',
      'v3.1.3 — Notes/Seller Notes를 다시 읽기 전용으로 변경 (실제 수정은 Zendesk/Seller Central 원본에서만)',
      'v3.1.4 — 사용되지 않는 코드 제거 (5141→5026줄), 기능 변화 없음',
      'v3.1.5 — Seller Notes 로딩 스피너 추가, Notes 첫 조회 시 안 보이던 문제 개선, 업데이트 안내 팝업에 v3.0.0~v3.1.5 요약 등록',
      'v3.1.6 — Notes가 안 보이던 근본 원인 수정: 텍스트 영역 자체가 아직 렌더링 안 된 상태였음 (값이 안 채워진 게 아니라)',
      'v3.1.7 — Customer context 탭 버튼 자체가 아직 렌더링 안 됐을 때 시도조차 안 하고 포기하던 문제 수정',
      'v3.1.8 — 사이드바 전체가 몇 초간 렌더링 안 되는 경우까지 고려해 대기 시간 연장 (2초→10초/5초)',
      'v3.2.0 — 내부 성능/구조 개선 6건 (기능 변화 없음): 페이지 텍스트 캐싱, 폴링 로직 통합, SC 동시 요청 수 제한, Auto-Fill 비동기 처리 방식(Promise.all) 개선, 일부 API 호출 통합',
      'v3.2.1 — 사이드바를 아예 못 불러왔을 때 "notes 없음" 대신 새로고침 안내 메시지 표시',
    ],
  };

  // One-click copy button appended to each value row (no drag + Cmd+C needed).
  const COPY_SVG = '<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/></svg>';
  const COPY_BTN = `<button class="sp-copy" title="Copy" tabindex="-1" aria-label="Copy">${COPY_SVG}</button>`;
  // Same ring spinner used everywhere in the panel — was previously an iOS-style
  // 12-dot spinner here but a plain rotating ring elsewhere (setStatus's
  // "Fetching order data…" spinner); unified to this one style everywhere.
  const SPINNER_HTML = `<span style="display:inline-block;width:12px;height:12px;border:2px solid rgba(110,110,115,0.25);border-top-color:#6e6e73;border-radius:50%;animation:sp-spin 0.8s linear infinite;vertical-align:middle;" role="img" aria-label="로딩중"></span>`;

  // ── Zendesk custom field IDs ─────────────────────────────────────────────
  const ZD = {
    ORDER_ID:      360021934132,
    ASIN:          360021934312,
    SKU:           900008676703,
    CUST_NAME:     360021999951,
    ORDER_STATUS:  360021934152,
    ORDER_TOTAL:   360021934172,
    DELIVERY_LVL:  900003828503,
    PURCHASE_DATE: 360019586172,
    COUNTRY:       4513936822297,
    FULFILLMENT:   900002781823,
    POINT_OF_PUR:  20016270875033,
    DEVICE:        360022185671,
    PRODUCT_NAME:  360022185891,
    BRAND_DETAIL:  5495572594201,
    PHOTO_EXIST:   26936618247577,
    TOTAL_ORDERS:  21714421937305,
    TOTAL_REFUNDS: 21745453864345,
    SPIGEN_REFUND: 21745465897369,
    DAEBUNRYU:     58529165213721,
    SAENGSAN:      58529176884761,
    ORIGIN_INFO:   58529167605273,
  };

  const COUNTRY_MAP = {
    US:'us', GB:'uk', DE:'de', FR:'fr', IT:'it', ES:'es', JP:'jp',
    NL:'nl', SE:'se', IE:'ie', PL:'pl', TR:'tr', BE:'be', IN:'in',
    SG:'sg', AU:'au', CA:'ca', MX:'mx', KR:'kr',
  };

  // Target for the "Detected language" auto-correct (see autoCorrectDetectedLanguage_
  // below) — keyed by the same lowercase COUNTRY_MAP values / Country* custom
  // field tokens, mapped to Zendesk's own exact language-picker option text
  // (captured live from the real dropdown's 113-item list). 'be' is
  // deliberately omitted: Belgium has no single dominant order language
  // (French/Dutch mixed), so guessing would be worse than leaving Zendesk's
  // own (possibly wrong) detection alone.
  const LANGUAGE_BY_COUNTRY_VALUE = {
    us:'English', uk:'English', ie:'English', au:'English', sg:'English', ca:'English', in:'English',
    de:'German', fr:'French', it:'Italian', jp:'Japanese', nl:'Dutch', se:'Swedish',
    pl:'Polish', tr:'Turkish', kr:'Korean',
    es:'Spanish (Spain)', mx:'Spanish (Latin America)',
  };

  const SCRIPT_VER = (typeof GM_info !== 'undefined' ? GM_info?.script?.version : null) || '3.6.2';
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  // Poll checkFn every intervalMs until it returns a truthy value or
  // maxTries is reached; resolves with whatever the last call returned
  // (falsy on timeout). checkFn must return a genuinely falsy value (null/
  // false/undefined/0) to mean "not ready yet" — an empty array/NodeList is
  // truthy in JS, so callers checking "did I find any elements" need to map
  // that explicitly (e.g. `() => arr.length ? arr : null`).
  function pollUntil_(checkFn, { intervalMs = 250, maxTries = 20 } = {}) {
    return new Promise(resolve => {
      let tries = 0;
      const tick = () => {
        const result = checkFn();
        if (result || ++tries >= maxTries) { resolve(result); return; }
        setTimeout(tick, intervalMs);
      };
      tick();
    });
  }

  // Promise wrapper around GM_xmlhttpRequest — same options in (method, url,
  // headers, data, timeout, redirect, ...), resolves with the exact same
  // response object GM_xmlhttpRequest's onload receives (status,
  // responseText, responseHeaders, finalUrl, ...). No behavior change from
  // raw GM_xmlhttpRequest, just callback-shape -> Promise-shape, so existing
  // call sites can migrate one at a time without any risk of subtly changing
  // what gets sent or how a response is interpreted. Deliberately only used
  // for straightforward same-origin Zendesk API GETs so far — the Seller
  // Central / cookie-timing / redirect-sensitive call sites elsewhere in this
  // file are left on raw GM_xmlhttpRequest, since those were carefully
  // debugged today and a mechanical rewrite risks reintroducing something.
  function gmRequest_(opts) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        ...opts,
        onload: resolve,
        onerror: () => reject(new Error('network')),
        ontimeout: () => reject(new Error('timeout')),
      });
    });
  }

  // ── Fast data path (v3.8.0) ──────────────────────────────────────────────
  // Order + product lookups go to the GCX Reply Cloudflare Worker first
  // (GAS_Zendesk/GCXReply_Worker — same JSON as GAS ?orderId= / ?asin=, ~0.3-1 s
  // vs GAS's 2-7 s platform overhead). GAS stays the fallback for everything:
  // no Worker key, Worker error / non-200 / {error} / HTML, or no answer within
  // WORKER_HEDGE_MS → the GAS request fires and whichever valid answer lands
  // first wins. Every caller keeps its own GM_xmlhttpRequest-style handlers
  // (onload/onerror/ontimeout), so the response-handling code is unchanged.
  // ABM relay / claims / inferReason / MCF still talk to GAS directly.
  //
  // The Worker key is NOT in this (public) file: it's read from a Zendesk
  // dynamic-content item that only signed-in agents can fetch, then kept in
  // GM storage. 401 from the Worker → key dropped and re-read next time.
  const WORKER_URL       = '__WORKER_URL__';
  const WORKER_KEY_DC_ID = '__DC_ITEM_ID__';
  const WORKER_HEDGE_MS  = 3500;
  const ORDER_CLIENT_TTL = 5 * 60 * 1000; // reopened ticket within 5 min → instant (Lookup button = fresh)
  const _backendCache    = new Map(); // query → { res, ts }
  const _backendInflight = new Map(); // query → [opts, ...]
  let _workerKeyFetching = false;

  function workerKey_() {
    let k = null;
    try { k = GM_getValue('gcx_worker_key', null); } catch (_) {}
    if (!k && !_workerKeyFetching && WORKER_KEY_DC_ID.indexOf('__') !== 0) {
      _workerKeyFetching = true;
      gmRequest_({ method: 'GET', url: `https://spigenhelp.zendesk.com/api/v2/dynamic_content/items/${WORKER_KEY_DC_ID}.json`, timeout: 10000 })
        .then(res => {
          if (res.status !== 200) return;
          const item = JSON.parse(res.responseText).item || {};
          const v = (item.variants || []).find(x => x.default) || (item.variants || [])[0];
          const key = v && String(v.content || '').trim();
          if (key && /^[A-Za-z0-9_-]{24,}$/.test(key)) GM_setValue('gcx_worker_key', key);
        })
        .catch(() => {})
        .finally(() => { _workerKeyFetching = false; });
    }
    return k;
  }

  function isGoodJsonRes_(res) {
    if (!res || res.status !== 200) return false;
    const t = res.responseText || '';
    if (!t || t.trimStart().startsWith('<')) return false;
    try { return !JSON.parse(t).error; } catch (_) { return false; }
  }

  // query: 'orderId=…' / 'asin=…'. opts: GM_xmlhttpRequest-style (timeout,
  // onload, onerror, ontimeout). o.cacheTtl: keep good answers client-side;
  // o.fresh: skip client cache + in-flight sharing and ask the Worker for a
  // fresh (non-cached) answer.
  function backendRequest_(query, opts, o) {
    o = o || {};
    const now = Date.now();
    if (!o.fresh) {
      const hit = _backendCache.get(query);
      if (hit && now - hit.ts < (o.cacheTtl || 0)) { setTimeout(() => opts.onload && opts.onload(hit.res), 0); return; }
      const waiting = _backendInflight.get(query);
      if (waiting) { waiting.push(opts); return; }
    }
    const waiters = [opts];
    if (!o.fresh) _backendInflight.set(query, waiters);

    let settled = false, gasFired = false, pending = 0, lastBad = null, hedgeT = null;
    const deliver = (kind, res) => {
      if (settled) return;
      settled = true;
      clearTimeout(hedgeT);
      if (!o.fresh && _backendInflight.get(query) === waiters) _backendInflight.delete(query);
      if (kind === 'load' && o.cacheTtl && isGoodJsonRes_(res)) _backendCache.set(query, { res, ts: Date.now() });
      waiters.forEach(w => {
        try {
          if (kind === 'load') w.onload && w.onload(res);
          else if (kind === 'timeout') w.ontimeout && w.ontimeout();
          else w.onerror && w.onerror(res);
        } catch (e) { console.error('[GCX] backend handler', e); }
      });
    };
    const maybeFinishBad = () => { if (!pending && !settled && gasFired) deliver(lastBad ? lastBad.kind : 'error', lastBad && lastBad.res); };

    const fireGas = () => {
      if (gasFired || settled) return;
      gasFired = true; pending++;
      GM_xmlhttpRequest({
        method: 'GET', url: `${GAS_URL}?${query}`, redirect: 'follow', timeout: opts.timeout || 15000,
        onload(res)   { pending--; if (isGoodJsonRes_(res) || !pending) deliver('load', res); else { lastBad = { kind: 'load', res }; } },
        onerror(e)    { pending--; lastBad = lastBad || { kind: 'error', res: e }; maybeFinishBad(); },
        ontimeout()   { pending--; lastBad = lastBad || { kind: 'timeout' }; maybeFinishBad(); },
      });
    };

    const key = WORKER_URL.indexOf('__') === 0 ? null : workerKey_();
    if (!key) { fireGas(); return; }
    pending++;
    GM_xmlhttpRequest({
      method: 'GET', url: `${WORKER_URL}/?${query}${o.fresh ? '&fresh=1' : ''}`,
      headers: { 'x-gcx-key': key }, timeout: Math.min(opts.timeout || 15000, 12000),
      onload(res) {
        pending--;
        if (isGoodJsonRes_(res)) { deliver('load', res); return; }
        if (res.status === 401) { try { GM_setValue('gcx_worker_key', null); } catch (_) {} }
        if (!gasFired) fireGas(); else if (lastBad && !pending) deliver(lastBad.kind, lastBad.res);
      },
      onerror()   { pending--; if (!gasFired) fireGas(); else maybeFinishBad(); },
      ontimeout() { pending--; if (!gasFired) fireGas(); else maybeFinishBad(); },
    });
    hedgeT = setTimeout(fireGas, WORKER_HEDGE_MS);
  }

  // ── Seller Central request sharing (v3.8.0) ──────────────────────────────
  // The same /orders-api/order/{id} response used to be fetched 2-3 times per
  // ticket (Seller SKU lookup, buyer stats, refund check), each only after GAS
  // returned. scRequest_ shares one in-flight request per URL and keeps clean
  // JSON answers for 60 s. Prefetch: as soon as a ticket's order ID is known,
  // the SC order call starts on the SC domain previously seen for that order-ID
  // prefix (learned from real results), in parallel with the backend lookup.
  // Wrong guess = one unused request; the real call path is unchanged.
  const SC_SHARE_TTL   = 60 * 1000;
  const _scShareCache  = new Map();
  const _scShareFlight = new Map();
  function scRequest_(opts) {
    const url = opts.url;
    const hit = _scShareCache.get(url);
    if (hit && Date.now() - hit.ts < SC_SHARE_TTL) { setTimeout(() => opts.onload && opts.onload(hit.res), 0); return; }
    const waiting = _scShareFlight.get(url);
    if (waiting) { waiting.push(opts); return; }
    const waiters = [opts];
    _scShareFlight.set(url, waiters);
    const done = (kind, res) => {
      if (_scShareFlight.get(url) === waiters) _scShareFlight.delete(url);
      waiters.forEach(w => {
        try {
          if (kind === 'load') w.onload && w.onload(res);
          else if (kind === 'timeout') w.ontimeout && w.ontimeout();
          else w.onerror && w.onerror(res);
        } catch (e) { console.error('[GCX] sc handler', e); }
      });
    };
    GM_xmlhttpRequest({
      method: 'GET', url, redirect: 'follow',
      headers: opts.headers || { 'Accept': 'application/json' },
      timeout: Math.max(...waiters.map(w => w.timeout || 15000)),
      onload(res) {
        const t = res.responseText || '';
        if (res.status === 200 && !res.finalUrl?.includes('/ap/signin') && !t.trimStart().startsWith('<')) {
          _scShareCache.set(url, { res, ts: Date.now() });
        }
        done('load', res);
      },
      onerror(e) { done('error', e); },
      ontimeout() { done('timeout'); },
    });
  }
  function scBaseFromResult_(orderId, salesChannel, countryCode) {
    const u = sellerCentralUrl(orderId, salesChannel, countryCode);
    return u ? u.match(/^https:\/\/[^/]+/)[0] : null;
  }
  function learnScBase_(orderId, base) {
    if (!orderId || !base) return;
    try {
      const m = GM_getValue('gcx_sc_base_by_prefix', {}) || {};
      const p = orderId.slice(0, 3);
      if (m[p] !== base) { m[p] = base; GM_setValue('gcx_sc_base_by_prefix', m); }
    } catch (_) {}
  }
  function prefetchScOrder_(orderId) {
    let base = null;
    try { base = (GM_getValue('gcx_sc_base_by_prefix', {}) || {})[orderId.slice(0, 3)] || null; } catch (_) {}
    if (base) scRequest_({ url: `${base}/orders-api/order/${orderId}`, timeout: 15000, onload() {}, onerror() {}, ontimeout() {} });
  }

  // Started right after a ticket switch (the panel's own render flow still
  // runs on its usual schedule — it simply finds these answers ready or in
  // flight). Only the Zendesk API is used here, never the page DOM, which may
  // still show the previous ticket at this point.
  function prefetchTicketData_() {
    if (!/\/tickets\/\d+/.test(location.pathname)) return;
    getTicketFields((orderId, asin) => {
      try {
        if (orderId) {
          backendRequest_(`orderId=${encodeURIComponent(orderId)}`, { timeout: 15000, onload() {}, onerror() {}, ontimeout() {} }, { cacheTtl: ORDER_CLIENT_TTL });
          prefetchScOrder_(orderId);
        }
        if (asin && getDataFetchPrefs().fetchProduct) {
          gasProductRequest_(asin, { method: 'GET', url: `${GAS_URL}?asin=${encodeURIComponent(asin)}`, redirect: 'follow', timeout: 30000, onload() {}, onerror() {} });
        }
      } catch (e) { console.error('[GCX] prefetch', e); }
    });
  }

  // ── Module state ─────────────────────────────────────────────────────────
  let lastOrderData    = null;
  let lastProductData  = null;
  let _panelSession    = 0; // incremented on every resetPanel(); guards stale async callbacks
  let _productReady    = false; // true once product lookup finishes (or is determined impossible)
  let _gcrFilledThisTicket = false; // true after Auto-Fill confirmed & submitted on current ticket
  let _panelEl         = null; // live reference to the panel (survives React detaching it when docked)

  // ── UI state persistence ──────────────────────────────────────────────────
  function loadUi() {
    try { return JSON.parse(localStorage.getItem('gcx_ui') || '{}'); } catch { return {}; }
  }
  function saveUi(patch) {
    try { localStorage.setItem('gcx_ui', JSON.stringify(Object.assign(loadUi(), patch))); } catch {}
  }
  // ABM relay alerts (auto toasts + the floating red badge) — on by default;
  // agents who find the constant popups noisy can turn them off in the ⚙
  // settings drawer and use the header's static "ABM" button (with its quiet
  // pending count) to open the relay-status panel on demand instead.
  function abmAlertsEnabled_() { return loadUi().abmAlertsOff !== true; }
  function applySectionState(container) {
    const c = loadUi().collapsed || {};
    container.querySelectorAll('[data-sp-section]').forEach(block => {
      const key = block.dataset.spSection;
      if (!(key in c)) return;
      if (c[key]) block.classList.add('collapsed');
      else block.classList.remove('collapsed');
    });
  }

  // ── Data fetch toggle preferences ──────────────────────────────────────────
  function getDataFetchPrefs() {
    try {
      const prefs = JSON.parse(localStorage.getItem('gcx_data_fetch_prefs') || '{}');
      return {
        fetchOrder: prefs.fetchOrder !== false,
        fetchShipping: prefs.fetchShipping !== false,
        fetchProduct: prefs.fetchProduct !== false,
      };
    } catch { return { fetchOrder: true, fetchShipping: true, fetchProduct: true }; }
  }
  function saveDataFetchPrefs(prefs) {
    try { localStorage.setItem('gcx_data_fetch_prefs', JSON.stringify(prefs)); } catch {}
  }
  let lastAmazonProduct = null;
  let lastAiReason      = null;

  // ── Ticket JSON cache: avoid double-fetching the same ticket JSON ────────
  // Both getTicketFields (auto-detect) and autoFillTicket (currentCfMap) call
  // /api/v2/tickets/{id}.json on the same ticket. This cache makes the second
  // caller reuse the first response (60s TTL — fresh enough for one session).
  let _ticketJsonCache = null; // { id, ticket, ts }

  // SC buyer-stats caches (in-memory, session-scoped)
  // _scEmailCache_[orderId]   → buyer marketplace email (immutable once known)
  // _scOrdersCache_[email]    → { count, orderIds, ts }  (15-min TTL; order list grows slowly)
  // _scRefundCache_[orderId]  → true|false  (immutable for completed orders)
  // _scOrderQtyCache_[orderId]→ total item quantity (sum of QuantityOrdered) for that order
  const _scEmailCache_    = {};
  const _scOrdersCache_   = {};
  const _scRefundCache_   = {};
  const _scOrderQtyCache_ = {};
  function fetchTicketJson_(ticketId, cb) {
    const now = Date.now();
    if (_ticketJsonCache && _ticketJsonCache.id === ticketId && now - _ticketJsonCache.ts < 60000) {
      cb(_ticketJsonCache.ticket);
      return;
    }
    gmRequest_({ method: 'GET', url: `https://spigenhelp.zendesk.com/api/v2/tickets/${ticketId}.json`, timeout: 12000 })
      .then(res => {
        if (res.status !== 200) { cb(null); return; }
        try {
          const ticket = JSON.parse(res.responseText).ticket || {};
          _ticketJsonCache = { id: ticketId, ticket, ts: Date.now() };
          cb(ticket);
        } catch { cb(null); }
      })
      .catch(() => cb(null));
  }

  // ── SC session expiry warning ─────────────────────────────────────────────
  // Shown when fetchScItems / fetchScBuyerStats_ detect a login redirect.
  // Inserts a yellow banner at the top of #sp-result; removed on ticket nav.
  function showScSessionWarning_(scBaseUrl) {
    if (document.getElementById('sp-sc-session-warn')) return;
    const result = document.getElementById('sp-result');
    if (!result) return;
    const domain = (scBaseUrl || 'https://sellercentral.amazon.de').match(/^https:\/\/[^/]+/)?.[0] || 'https://sellercentral.amazon.de';
    const warn = document.createElement('div');
    warn.id = 'sp-sc-session-warn';
    warn.style.cssText = 'background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:6px 10px;font-size:11.5px;color:#856404;margin-bottom:8px;line-height:1.4;';
    warn.innerHTML = `⚠ SC 세션 만료 — <a href="${domain}/orders-v3" target="_blank" rel="noopener" style="color:#856404;font-weight:600;text-decoration:underline;">Seller Central 로그인</a> 필요 (구매이력/SKU 없음)`;
    result.insertBefore(warn, result.firstChild);
  }

  // ── Zendesk API: read order ID + ASIN from ticket custom fields ──────────
  function getTicketFields(cb) {
    const m = location.pathname.match(/\/tickets\/(\d+)/);
    if (!m) return cb(null, null, []);
    fetchTicketJson_(m[1], ticket => {
      if (!ticket) return cb(null, null, []);
      const fields  = ticket.custom_fields || [];
      const vals    = fields.map(f => String(f.value || ''));
      const orderId = vals.find(v => /^\d{3}-\d{7}-\d{7}$/.test(v)) || null;
      const asin    = vals.find(v => /^B[A-Z0-9]{9}$/.test(v)) || null;
      const desc    = ticket.description || '';
      const bodyIds = [...new Set([...desc.matchAll(/\b(\d{3}-\d{7}-\d{7})\b/g)].map(x => x[1]))];
      cb(orderId, asin, bodyIds);
    });
  }

  // ── Auto-fill helpers ────────────────────────────────────────────────────

  function salesChannelToPOP(ch) {
    if (!ch) return null;
    const s = ch.toLowerCase();
    if (s.includes('.co.uk'))  return 'amazon_united_kingdom';
    if (s.includes('.co.jp'))  return 'amazon_japan';
    if (s.includes('.com.sg')) return 'amazon_singapore';
    if (s.includes('.in'))     return 'amazon_india';
    if (s.includes('.de') || s.includes('.fr') || s.includes('.it') ||
        s.includes('.es') || s.includes('.nl')) return 'amazon_eu';
    return 'others';
  }

  const COUNTRY_SC = {
    DE:'amazon.de', FR:'amazon.fr', IT:'amazon.it', ES:'amazon.es',
    NL:'amazon.nl', PL:'amazon.pl', SE:'amazon.se', BE:'amazon.com.be',
    IE:'amazon.ie', GB:'amazon.co.uk', JP:'amazon.co.jp', IN:'amazon.in',
    SG:'amazon.com.sg', AU:'amazon.com.au', CA:'amazon.ca',
    MX:'amazon.com.mx', TR:'amazon.com.tr', US:'amazon.com',
  };

  // EU marketplaces (non-DE) all share order search AND ABM case messaging on
  // DE Seller Central. Logging into DE SC is sufficient — no need to log into
  // each marketplace SC. amazon.ie added 2026-08-06: ABM_Relay_Log showed
  // persistent "case lookup failed — no SC login session for amazon.ie" with
  // high Attempts counts (e.g. rows 404/520/535) because this set omitted it,
  // so abmDomainFromAddress_/scDomain_ kept resolving IE cases to
  // sellercentral.amazon.ie, which no agent is ever logged into — Ireland's
  // ABM cases actually live on the same DE Seller Central messaging page as
  // the rest of EU, confirmed by the user.
  // amazon.com.be fixed 2026-08-14: same bug, Belgium this time. The set had
  // 'amazon.be' — not a real Amazon storefront domain at all (Belgium's is
  // amazon.com.be, matching the buyer-proxy address's actual
  // ...@marketplace.amazon.com.be) — so it never matched and every BE case
  // kept resolving to sellercentral.amazon.com.be, again a domain no agent
  // is logged into. Confirmed live: ABM_Relay_Log row 584, 44 failed
  // attempts, "no SC login session for amazon.com.be". Also fixed the same
  // wrong value in COUNTRY_SC.BE above (the salesChannel-absent fallback
  // path) for consistency, even though the address-based path above is what
  // actually hit this bug.
  const EU_SC_REDIRECT = new Set(['amazon.fr','amazon.it','amazon.es','amazon.nl','amazon.pl','amazon.se','amazon.com.be','amazon.ie']);

  function scDomain_(salesChannel, countryCode) {
    const domain = salesChannel ? salesChannel.toLowerCase()
      : (countryCode ? (COUNTRY_SC[countryCode] || null) : null);
    if (!domain) return null;
    return EU_SC_REDIRECT.has(domain) ? 'amazon.de' : domain;
  }

  // Every other marketplace's Seller Central follows the same
  // https://sellercentral.{marketplace-domain} pattern (sellercentral.amazon.de,
  // sellercentral.amazon.co.uk, ...) — Japan does not. Its real Seller Central
  // is sellercentral-japan.amazon.com, confirmed by the user. sellercentral.
  // amazon.co.jp DOES resolve to a real Amazon page (an account-switcher
  // screen), which is exactly why this went unnoticed for a while — it looks
  // like a working domain, but it's an account-switcher-gated path most
  // agents' sessions were never actually authenticated against, not the real
  // JP Seller Central agents use day to day (that's what getMcfBase_ already
  // used correctly for the MCF flow — this brings every OTHER sellercentral.
  // ${domain} builder in this file, including the SC-login-check banner that
  // was persistently telling a logged-in agent "JP 로그인이 필요합니다", in
  // line with it). Centralized here instead of special-casing JP at each of
  // the ~8 call sites individually.
  function scBaseUrl_(domain) {
    if (!domain) return null;
    if (domain === 'amazon.co.jp') return 'https://sellercentral-japan.amazon.com';
    return `https://sellercentral.${domain}`;
  }

  function sellerCentralUrl(orderId, salesChannel, countryCode) {
    if (!orderId) return null;
    const base = scBaseUrl_(scDomain_(salesChannel, countryCode));
    return base ? `${base}/orders-v3/order/${orderId}` : null;
  }

  // Build Seller Central buyer order history search URL (last 2 years)
  function sellerCentralSearchUrl_(salesChannel, countryCode, buyerEmail) {
    if (!buyerEmail) return null;
    const base = scBaseUrl_(scDomain_(salesChannel, countryCode));
    if (!base) return null;
    const now         = Date.now();
    const twoYearsAgo = Math.round(now - 2 * 365.25 * 24 * 3600 * 1000);
    return `${base}/orders-v3/search?qt=email&q=${encodeURIComponent(buyerEmail)}&date-range=${twoYearsAgo}-${now}`;
  }

  // Derive amazon.XX domain from order SalesChannel / CountryCode
  function amazonDomain_(salesChannel, countryCode) {
    if (salesChannel) return salesChannel.toLowerCase(); // "amazon.in", "amazon.co.jp", etc.
    return countryCode ? (COUNTRY_SC[countryCode] || 'amazon.com') : 'amazon.com';
  }

  // Try to extract Spigen model name from product title
  // "Spigen Liquid Air | iPhone 15 Case"   → "Liquid Air"
  // "Spigen Case for iPhone 15 Liquid Air Case..." → "Liquid Air"
  function modelFromTitle_(title) {
    if (!title) return '';
    // Pattern 1: "Brand Model | Device ..." — model between brand and pipe
    const pipeMatch = title.match(/^[A-Za-z]+\s+(.+?)\s+[|｜]/);
    if (pipeMatch && !/^(case|cover|protector|glass|film)/i.test(pipeMatch[1]))
      return pipeMatch[1].trim();
    // Pattern 2: strip "Spigen (Case|Cover) for" prefix, then strip device name, grab until next Case/Cover
    const noBrand = title.replace(/^[A-Za-z]+\s+/i, '').replace(/^(?:Case|Cover)\s+for\s+/i, '');
    const devStrip = noBrand.replace(
      /^(?:for\s+)?(?:iPhone|Samsung(?:\s+Galaxy)?|Galaxy|Google\s+Pixel|Pixel|iPad|Huawei|Xiaomi|OnePlus|LG)\s+[\w\s]*?(?=\b[A-Z][a-z])/,
      ''
    );
    const modelMatch = devStrip.match(/^([A-Z][^\|]+?)\s+(?:Case|Cover|Protector|Glass|Film|Screen)\b/i);
    if (modelMatch) return modelMatch[1].trim();
    return '';
  }

  // Parse Amazon product page static HTML → sheet-column-shaped object.
  // Handles both tech-spec table (amazon.in/com) and po-* overview rows (amazon.de etc.)
  function parseAmazonPage_(doc) {
    const spec = {};

    // Tech spec table — amazon.com, amazon.in, amazon.co.jp (with ?language=en_GB)
    doc.querySelectorAll(
      '#productDetails_techSpec_section_1 tr, #productDetails_db_sections tr, .prodDetTable tr'
    ).forEach(tr => {
      const k = tr.querySelector('th')?.textContent?.trim();
      const v = tr.querySelector('td')?.textContent?.replace(/\s+/g, ' ').trim();
      if (k && v) spec[k] = v;
    });

    // Product overview rows — amazon.de, amazon.fr, etc. (after ?language=en_GB → English keys)
    doc.querySelectorAll('tr[class*="po-"]').forEach(row => {
      const k = row.querySelector('.a-span3 span, td:first-child span')?.textContent?.trim();
      const v = row.querySelector('.a-span9 span, td:last-child span')?.textContent?.trim();
      if (k && v) spec[k] = v;
    });

    const title = doc.querySelector('#productTitle')?.textContent?.trim() || '';
    const mfr   = (spec['Manufacturer'] || '').split('/')[0].replace(/,.*$/, '').trim();

    return {
      _title:   title,   // full Amazon product page title (#productTitle)
      SKU:      spec['Model Number']                                              || '',
      '모델명':  spec['Model Name']  || modelFromTitle_(title)                   || '',
      '브랜드':  spec['Brand Name']  || spec['Brand']                            || 'Spigen',
      '제조사명': mfr,
      '기종명':  spec['Compatible Phone Models'] || spec['Compatible phone models'] ||
                 spec['Compatible Devices']      || spec['Compatible devices']      || '',
      '색상명':  spec['Colour'] || spec['Color']                                 || '',
      '대분류':  spec['Form Factor'] || spec['Item Type Name']                   || '',
      '생산업체': mfr,
      '원산지정보': spec['Country of Origin'] || spec['Country of origin']       || '',
    };
  }

  // Fetch amazon.XX/dp/{asin} HTML and parse product info; cb(product|null, pageUrl)
  // Falls back through: order's marketplace → amazon.co.jp → amazon.com
  //
  // Memoized per ASIN+domain: the /dp/ page is 1-2 MB and takes 4-7 s, and it
  // used to be re-downloaded on every ticket (Product Info fallback AND the
  // ASIN Sources block each fetched it). Concurrent callers share one in-flight
  // request; successful results persist 12 h in GM storage (capped at
  // AMZ_CACHE_MAX entries) so reopening a ticket / another ticket with the
  // same ASIN is instant. Not-found results are never cached (could be a
  // transient captcha/503).
  const AMZ_CACHE_KEY = 'gcx_amz_product_cache';
  const AMZ_CACHE_TTL = 12 * 3600 * 1000;
  const AMZ_CACHE_MAX = 300;
  const _amzInflight_ = {};
  function amzCacheRead_(key) {
    try {
      const e = (GM_getValue(AMZ_CACHE_KEY, {}) || {})[key];
      return e && Date.now() - e.ts < AMZ_CACHE_TTL ? e : null;
    } catch (_) { return null; }
  }
  function amzCacheWrite_(key, product, url) {
    try {
      const all = GM_getValue(AMZ_CACHE_KEY, {}) || {};
      all[key] = { product, url, ts: Date.now() };
      const keys = Object.keys(all);
      if (keys.length > AMZ_CACHE_MAX) {
        keys.sort((a, b) => all[a].ts - all[b].ts).slice(0, keys.length - AMZ_CACHE_MAX).forEach(k => delete all[k]);
      }
      GM_setValue(AMZ_CACHE_KEY, all);
    } catch (_) {}
  }
  function fetchAmazonProduct_(asin, cb) {
    const primaryDomain = amazonDomain_(lastOrderData?.order?.SalesChannel, lastOrderData?.address?.CountryCode);
    const cacheKey = `${asin}|${primaryDomain}`;
    const cached = amzCacheRead_(cacheKey);
    if (cached) { cb(cached.product, cached.url); return; }
    if (_amzInflight_[cacheKey]) { _amzInflight_[cacheKey].push(cb); return; }
    _amzInflight_[cacheKey] = [cb];
    const cb_ = cb;
    cb = (product, url) => {
      const cbs = _amzInflight_[cacheKey] || [cb_];
      delete _amzInflight_[cacheKey];
      if (product) amzCacheWrite_(cacheKey, product, url);
      cbs.forEach(f => { try { f(product, url); } catch (e) { console.error('[GCX] amazon cb', e); } });
    };
    const primaryUrl    = `https://www.${primaryDomain}/dp/${asin}`;

    const fallbacks = [primaryDomain];
    if (!fallbacks.includes('amazon.co.jp')) fallbacks.push('amazon.co.jp');
    if (!fallbacks.includes('amazon.com'))   fallbacks.push('amazon.com');

    function tryDomain(idx) {
      if (idx >= fallbacks.length) return cb(null, primaryUrl);
      const domain    = fallbacks[idx];
      const langParam = domain === 'amazon.com' ? '' : '?language=en_GB';
      const url       = `https://www.${domain}/dp/${asin}${langParam}`;
      GM_xmlhttpRequest({
        method:   'GET',
        url,
        headers:  { 'Accept-Language': 'en-GB,en;q=0.9', 'Accept': 'text/html' },
        redirect: 'follow',
        timeout:  20000,
        onload(res) {
          if (res.status !== 200) return tryDomain(idx + 1);
          try {
            const doc     = new DOMParser().parseFromString(res.responseText, 'text/html');
            const product = parseAmazonPage_(doc);
            if (Object.values(product).some(v => v)) return cb(product, url);
            tryDomain(idx + 1);
          } catch { tryDomain(idx + 1); }
        },
        onerror()   { tryDomain(idx + 1); },
        ontimeout() { tryDomain(idx + 1); },
      });
    }

    tryDomain(0);
  }

  // Normalize label text: strip ★ * ( ) . and trim, lowercase
  function normLabel(s) {
    return s.replace(/[^가-힣a-zA-Z0-9\s]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  // Fill a Zendesk React-controlled text/date input by label text
  function fillZdInput(labelText, value) {
    if (!value) return false;
    const needle = normLabel(labelText);
    for (const input of document.querySelectorAll(
      '[data-test-id="ticket-fields-text-field"], [data-test-id="ticket-fields-date-field"]'
    )) {
      let node = input.parentElement;
      for (let i = 0; i < 8 && node; i++, node = node.parentElement) {
        const lbl = node.querySelector('label');
        if (lbl && normLabel(lbl.textContent).startsWith(needle)) {
          const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
          setter.call(input, value);
          input.dispatchEvent(new Event('input', { bubbles: true }));
          return true;
        }
      }
    }
    return false;
  }

  // All Zendesk text/date field labels that GCX Reply ever writes via fillZdInput.
  // Used by clearAllZdFields_() to wipe React state on ticket navigation so stale
  // values from a previous ticket never get submitted when the agent marks Solved/Pending.
  const ZD_TEXT_FIELD_LABELS = [
    'Order ID', 'ASIN', '문의SKU', 'Customer Full Name', 'Purchase Date',
    'Order Status', 'Order Total', 'Delivery Level', '대분류', '생산업체', '원산지정보',
  ];

  function clearAllZdFields_(labels) {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    const needle_list = (labels || ZD_TEXT_FIELD_LABELS).map(normLabel);
    for (const input of document.querySelectorAll(
      '[data-test-id="ticket-fields-text-field"], [data-test-id="ticket-fields-date-field"]'
    )) {
      if (!input.value) continue; // already empty — skip
      let node = input.parentElement;
      for (let i = 0; i < 8 && node; i++, node = node.parentElement) {
        const lbl = node.querySelector('label');
        if (lbl) {
          const t = normLabel(lbl.textContent);
          if (needle_list.some(n => t.startsWith(n))) {
            setter.call(input, '');
            input.dispatchEvent(new Event('input', { bubbles: true }));
            break;
          }
        }
      }
    }
  }

  // Read the current value of a Zendesk React text/date input by label text
  function readZdInput_(labelText) {
    const needle = normLabel(labelText);
    for (const input of document.querySelectorAll(
      '[data-test-id="ticket-fields-text-field"], [data-test-id="ticket-fields-date-field"]'
    )) {
      let node = input.parentElement;
      for (let i = 0; i < 8 && node; i++, node = node.parentElement) {
        const lbl = node.querySelector('label');
        if (lbl && normLabel(lbl.textContent).startsWith(needle)) return input.value || '';
      }
    }
    return '';
  }

  // Returns the actual DOM label text for the ZD field whose label starts with `labelPrefix`.
  function resolveZdLabel_(labelPrefix) {
    const needle = normLabel(labelPrefix);
    for (const input of document.querySelectorAll(
      '[data-test-id="ticket-fields-text-field"], [data-test-id="ticket-fields-date-field"]'
    )) {
      let node = input.parentElement;
      for (let i = 0; i < 8 && node; i++, node = node.parentElement) {
        const lbl = node.querySelector('label');
        if (lbl && normLabel(lbl.textContent).startsWith(needle)) return lbl.textContent.trim();
      }
    }
    return null;
  }

  // Fetch Zendesk field options (for Device / Product Name matching)
  function fetchZdFieldOpts(fieldId, cb) {
    gmRequest_({ method: 'GET', url: `https://spigenhelp.zendesk.com/api/v2/ticket_fields/${fieldId}.json`, timeout: 12000 })
      .then(res => {
        try { cb(JSON.parse(res.responseText).ticket_field?.custom_field_options || []); }
        catch { cb([]); }
      })
      .catch(() => cb([]));
  }

  // Cached wrapper around fetchZdFieldOpts — field options rarely change per session.
  // Uses sessionStorage with a 30-min TTL to avoid repeated API calls on the same ticket.
  const ZD_OPTS_CACHE_TTL = 30 * 60 * 1000;
  function fetchZdFieldOptsCached(fieldId, cb) {
    const key = `sp_zdopts_${fieldId}`;
    try {
      const raw = sessionStorage.getItem(key);
      if (raw) {
        const { ts, opts } = JSON.parse(raw);
        if (Date.now() - ts < ZD_OPTS_CACHE_TTL) { cb(opts); return; }
      }
    } catch {}
    fetchZdFieldOpts(fieldId, opts => {
      try { sessionStorage.setItem(key, JSON.stringify({ ts: Date.now(), opts })); } catch {}
      cb(opts);
    });
  }

  // Tokenize: split camelCase + letter↔digit boundaries, lowercase, strip dots/specials.
  // Letter↔digit split ensures "Fold7"→"Fold 7" and "Flip7"→"Flip 7" match dropdown options.
  function tokenize_(s) {
    return s.replace(/([a-z])([A-Z])/g, '$1 $2')      // camelCase: "MagFit" → "Mag Fit"
            .replace(/([a-zA-Z])(\d)/g, '$1 $2')       // letter→digit: "Fold7" → "Fold 7"
            .replace(/(\d)([a-zA-Z])/g, '$1 $2')       // digit→letter: "7Pro" → "7 Pro"
            .toLowerCase().replace(/\./g, '').replace(/[^a-z0-9가-힣]+/g, ' ').trim().split(/\s+/).filter(Boolean);
  }

  // Jaccard similarity between two token arrays (set-based)
  function jaccard_(a, b) {
    const sa = new Set(a), sb = new Set(b);
    let inter = 0;
    for (const t of sa) if (sb.has(t)) inter++;
    const union = sa.size + sb.size - inter;
    return union ? inter / union : 0;
  }

  // Return top Product Name candidates when multiple close matches exist.
  // Returns [] (no match), [single] (clear winner), or [multi] (ambiguous → show picker).
  // Each element: { val, name (full), displayName (prefix-stripped), score }
  function topCandidateOpts_(opts, label, stripPrefix) {
    if (!label || !opts.length) return [];
    const labelToks = tokenize_(label);
    const scored = opts
      .map(o => {
        const displayName = stripPrefix ? o.name.replace(/^[^_]+_/, '') : o.name;
        return { val: o.value, name: o.name, displayName, score: jaccard_(tokenize_(displayName), labelToks) };
      })
      .filter(x => x.score >= 0.35)
      .sort((a, b) => b.score - a.score);
    if (!scored.length) return [];
    const best = scored[0];
    const second = scored[1];
    if (!second || best.score - second.score >= 0.25 || best.score >= 0.85) {
      return best.score >= 0.5 ? [best] : [];
    }
    return scored.filter(c => c.score >= best.score - 0.25).slice(0, 4);
  }

  // Return top Device candidates (ambiguous → show picker).
  // Each element: { val, name, displayName, score }
  function topDeviceCandidates_(opts, label, ticketText) {
    if (!label || !opts.length) return [];
    const labelToks = tokenize_(label.replace(/\//g, ' '));
    let maxScore = 0;
    const scored = opts.map(o => {
      const score = jaccard_(tokenize_(o.name), labelToks);
      if (score > maxScore) maxScore = score;
      return { val: o.value, name: o.name, displayName: o.name, score };
    });
    if (maxScore < 0.25) return [];
    const top = scored.filter(c => c.score >= maxScore * 0.95).sort((a, b) => b.score - a.score);
    if (top.length === 1) return top;
    if (ticketText) {
      const t = ticketText.toLowerCase();
      const mentioned = top.filter(c => {
        const base = c.name.toLowerCase().replace(/\s*\(.*?\)\s*/g, '').trim();
        return base.length > 4 && t.includes(base);
      });
      if (mentioned.length === 1) return [mentioned[0]];
    }
    top.sort((a, b) => {
      const aVer = parseInt((a.name.match(/\d+/) || ['0'])[0]);
      const bVer = parseInt((b.name.match(/\d+/) || ['0'])[0]);
      return bVer - aVer;
    });
    return top.slice(0, 4);
  }

  // Map 대분류 value → Brand(상세) tagger tag, SP/CASE first
  function brandFromDaebunryu(d) {
    if (!d) return null;
    if (d.includes('보호필름'))                  return 'spigen_sp_';
    if (d === '휴대폰케이스')                    return 'spigen_case_';
    if (d.includes('차량'))                     return 'spigen_new_biz_';
    if (/래저|음향|워치|주변기기|거치대/.test(d)) return 'spigen_sda_';
    return null;
  }

  // Fetch ticket comments; cb(true) if any customer comment has image/video attachment
  function fetchTicketComments(ticketId, cb) {
    gmRequest_({ method: 'GET', url: `https://spigenhelp.zendesk.com/api/v2/tickets/${ticketId}/comments.json?include=users`, timeout: 12000 })
      .then(res => {
        if (res.status !== 200) return cb(false);
        try {
          const data     = JSON.parse(res.responseText);
          const comments = data.comments || [];
          const users    = data.users    || [];
          const agentIds = new Set(users.filter(u => u.role !== 'end-user').map(u => u.id));
          const hasPhoto = comments.some(c =>
            !agentIds.has(c.author_id) &&
            (c.attachments || []).some(a => /^(image|video)\//.test(a.content_type))
          );
          cb(hasPhoto);
        } catch { cb(false); }
      })
      .catch(() => cb(false));
  }

  // ── Fill confirmation modal ──────────────────────────────────────────────

  function showFillConfirm_(rows, onConfirm, onCancel) {
    const overlay = document.createElement('div');
    overlay.id = 'sp-fill-confirm-overlay';
    overlay.style.cssText = [
      'position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:2147483646',
      'display:flex;align-items:center;justify-content:center;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
    ].join(';');

    const fillable = rows.filter(r => r.after || r.candidates?.length >= 2);
    const hasOpts_ = r => Array.isArray(r.opts) && r.opts.length > 0;
    const comboOpts = {}; // rowIdx -> full {value,name}[] list, for filter-as-you-type
    const card = document.createElement('div');
    card.style.cssText = [
      'background:#fff;border-radius:8px;padding:20px 24px 16px',
      'max-width:680px;width:94vw;max-height:82vh;overflow-y:auto',
      'box-shadow:0 8px 40px rgba(0,0,0,0.28);display:flex;flex-direction:column;gap:0',
    ].join(';');

    function esc_(s) { return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

    card.innerHTML = `
      <style>
        /* Unchecked "nothing to change" rows are dimmed, but the dimming must
           not carry into an open combobox suggestion list or make an input
           unreadable while the agent is actively editing it — :focus-within
           restores full opacity for the whole row the moment any input in it
           is focused (also fixes the combo suggestion list rendering
           semi-transparent and bleeding into the row underneath it). */
        .sp-confirm-dim { opacity: .5; }
        .sp-confirm-dim:focus-within { opacity: 1; }
      </style>
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;">
        <span style="font-weight:700;font-size:13px;color:#1f2d3d;">Confirm Auto-Fill</span>
        <label style="display:flex;align-items:center;gap:5px;font-size:11px;color:#555;cursor:pointer;user-select:none;">
          <input type="checkbox" id="sp-chk-all" style="margin:0;cursor:pointer;" ${fillable.length ? 'checked' : ''}> Select all
        </label>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:11.5px;">
        <thead>
          <tr style="border-bottom:2px solid #e4e8ed;">
            <th style="width:22px;padding:4px 6px;"></th>
            <th style="text-align:left;padding:4px 8px;color:#888;font-weight:600;white-space:nowrap;">Field</th>
            <th style="text-align:left;padding:4px 8px;color:#888;font-weight:600;min-width:130px;">Before</th>
            <th style="text-align:left;padding:4px 8px;color:#888;font-weight:600;min-width:160px;">After</th>
          </tr>
        </thead>
        <tbody id="sp-confirm-tbody">
          ${rows.map((r, i) => {
            const isMulti   = r.candidates?.length >= 2;
            const noChange  = !isMulti && r.before && r.before === r.after;
            const hasValue  = r.after || isMulti;
            // Every row is editable: dropdown for tagger/select-backed ZD fields
            // (full real option list, not just top-ranked matches), plain text
            // input for free-text fields — mirrors the ticket's own field type
            // so the agent can correct a wrong/missing auto-fill before submitting.
            const selectOpts = isMulti ? r.candidates.map(c => ({ value: c.val, name: c.displayName })) : r.opts;
            let afterTd;
            if (hasOpts_(r) || isMulti) {
              comboOpts[i] = selectOpts;
              const currentName = r.api?.value ? (selectOpts.find(o => o.value === r.api.value)?.name || '') : '';
              // Searchable combobox — type-to-filter, same as Zendesk's own
              // ticket-field dropdowns, instead of a plain <select> that only
              // filters by first-letter-jump. Suggestion list is a positioned
              // <div>, not a native <select>, so it can be filtered live.
              afterTd = `<td style="padding:3px 8px;position:relative;">
                   <input type="text" data-combo="${i}" data-value="${esc_(r.api?.value || '')}"
                     value="${esc_(currentName)}" placeholder="Type to search…" autocomplete="off"
                     style="font-size:11px;width:100%;max-width:190px;border:1px solid #b2d4be;border-radius:4px;padding:2px 5px;color:#1a6e3a;font-weight:600;background:#f4fbf7;box-sizing:border-box;">
                   <div data-combo-list="${i}" style="display:none;position:absolute;left:8px;right:8px;z-index:20;background:#fff;border:1px solid #d8dcde;border-radius:4px;max-height:160px;overflow-y:auto;box-shadow:0 4px 12px rgba(0,0,0,.18);margin-top:2px;"></div>
                 </td>`;
            } else if (r.dom?.isDate) {
              // A native <input type="date"> displays/types in whatever format
              // the OS locale dictates (dd/mm/yyyy, mm/dd/yyyy, …) — the
              // lang="sv-SE" trick to force ISO display turned out unreliable
              // across Chrome versions/OS locale settings in practice. Instead:
              // a plain text input that always shows/types literal yyyy-mm-dd
              // (our own format, not the browser's), paired with a hidden
              // native date input used ONLY for the calendar-pick interaction
              // — clicking the text field opens it via .showPicker(), and its
              // .value (always yyyy-mm-dd regardless of display locale) gets
              // mirrored back into the visible text field on change.
              afterTd = `<td style="padding:3px 8px;position:relative;">
                   <input type="text" data-textinput="${i}" data-datefield="${i}" value="${esc_(r.after)}"
                     placeholder="yyyy-mm-dd" pattern="\\d{4}-\\d{2}-\\d{2}"
                     style="font-size:11px;width:100%;max-width:190px;border:1px solid #d8dcde;border-radius:4px;padding:3px 6px;color:#1a6e3a;font-weight:600;background:#f4fbf7;box-sizing:border-box;">
                   <input type="date" data-datepicker="${i}" value="${esc_(r.after)}" tabindex="-1" aria-hidden="true"
                     style="position:absolute;left:0;bottom:0;opacity:0;width:1px;height:1px;border:0;padding:0;margin:0;pointer-events:none;">
                 </td>`;
            } else {
              afterTd = `<td style="padding:3px 8px;">
                   <input type="text" data-textinput="${i}" value="${esc_(r.after)}" placeholder="—"
                     style="font-size:11px;width:100%;max-width:190px;border:1px solid #d8dcde;border-radius:4px;padding:3px 6px;color:#1a6e3a;font-weight:600;background:#f4fbf7;box-sizing:border-box;">
                 </td>`;
            }
            return `<tr data-idx="${i}" class="${noChange ? 'sp-confirm-dim' : ''}" style="border-bottom:1px solid #f2f4f7;">
              <td style="padding:5px 6px;text-align:center;">
                <input type="checkbox" data-row="${i}" style="margin:0;cursor:pointer;"
                  ${hasValue && !noChange ? 'checked' : ''}>
              </td>
              <td style="padding:5px 8px;font-weight:500;color:#1f2d3d;white-space:nowrap;">${esc_(r.label)}</td>
              <td style="padding:5px 8px;color:#999;max-width:170px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"
                  title="${esc_(r.before)}">${esc_(r.before) || '<span style="color:#ccc">—</span>'}</td>
              ${afterTd}
            </tr>`;
          }).join('')}
        </tbody>
      </table>
      <div style="display:flex;align-items:center;gap:8px;justify-content:flex-end;margin-top:14px;padding-top:12px;border-top:1px solid #eef0f3;">
        <span id="sp-confirm-count" style="font-size:11px;color:#888;margin-right:auto;"></span>
        <button id="sp-confirm-cancel" style="padding:6px 16px;border:1px solid #d0d5dd;border-radius:5px;background:#fff;cursor:pointer;font-size:12px;color:#444;">Cancel</button>
        <button id="sp-confirm-ok" style="padding:6px 16px;border:none;border-radius:5px;background:#2a7a50;color:#fff;cursor:pointer;font-size:12px;font-weight:600;">Fill Selected</button>
      </div>`;

    overlay.appendChild(card);
    document.body.appendChild(overlay);

    const allChk  = card.querySelector('#sp-chk-all');
    const countEl = card.querySelector('#sp-confirm-count');
    const rowChks = () => [...card.querySelectorAll('input[data-row]:not([disabled])')]

    function refreshCount() {
      const n = rowChks().filter(c => c.checked).length;
      countEl.textContent = `${n} field${n !== 1 ? 's' : ''} selected`;
      card.querySelector('#sp-confirm-ok').disabled = n === 0;
    }
    refreshCount();

    allChk.addEventListener('change', () => {
      rowChks().forEach(c => c.checked = allChk.checked);
      refreshCount();
    });
    card.querySelector('#sp-confirm-tbody').addEventListener('change', e => {
      if (e.target.matches('input[data-row]')) {
        const all = rowChks();
        const checkedCount = all.filter(c => c.checked).length;
        allChk.indeterminate = checkedCount > 0 && checkedCount < all.length;
        allChk.checked = checkedCount === all.length;
        refreshCount();
        return;
      }
      // Editing the After text-input implies the agent wants it included
      // (combobox rows auto-check on item click instead — see below).
      if (e.target.matches('input[data-textinput]')) {
        const idx = +e.target.dataset.textinput;
        const chk = card.querySelector(`input[data-row="${idx}"]`);
        if (chk && !chk.checked) { chk.checked = true; refreshCount(); }
      }
    });
    // Date input (Purchase Date): clicking the visible yyyy-mm-dd text field
    // opens the hidden native date input's calendar instead — clicking the
    // text field directly can't show a calendar of its own, so the click is
    // redirected to the paired hidden <input type="date">.
    card.querySelector('#sp-confirm-tbody').addEventListener('click', e => {
      if (!e.target.matches('input[data-datefield]')) return;
      const idx = e.target.dataset.datefield;
      const picker = card.querySelector(`input[data-datepicker="${idx}"]`);
      if (picker) { try { picker.showPicker(); } catch (_) {} }
    });
    // Picking a date writes it back into the visible text field as yyyy-mm-dd
    // (a date input's .value is always ISO regardless of display locale) and
    // fires a real 'change' event so the existing data-textinput handler
    // (auto-check the row) runs — no separate logic needed here.
    card.querySelector('#sp-confirm-tbody').addEventListener('change', e => {
      if (!e.target.matches('input[data-datepicker]')) return;
      const idx = e.target.dataset.datepicker;
      const textEl = card.querySelector(`input[data-textinput="${idx}"]`);
      if (textEl && e.target.value) {
        textEl.value = e.target.value;
        textEl.dispatchEvent(new Event('change', { bubbles: true }));
      }
    });

    // ── Searchable combobox (type-to-filter dropdown) ─────────────────────
    // Replaces the old plain <select> for tagger/dropdown-backed fields —
    // matches Zendesk's own ticket-field dropdown behavior (type "iphone",
    // see every option containing "iphone", not just first-letter jump).
    const filterCombo_ = (idx, query) => {
      const opts = comboOpts[idx] || [];
      const q = query.trim().toLowerCase();
      const matches = q ? opts.filter(o => o.name.toLowerCase().includes(q)) : opts;
      return matches.slice(0, 50); // cap rendered suggestions
    };
    const renderComboList_ = (idx, query) => {
      const listEl = card.querySelector(`[data-combo-list="${idx}"]`);
      if (!listEl) return;
      const matches = filterCombo_(idx, query);
      listEl.innerHTML = matches.length
        ? matches.map(o => `<div class="sp-combo-item" data-combo-item="${idx}" data-val="${esc_(o.value)}" style="padding:5px 8px;cursor:pointer;font-size:11px;color:#1f2d3d;">${esc_(o.name)}</div>`).join('')
        : `<div style="padding:6px 8px;color:#999;font-size:11px;">No matches</div>`;
      listEl.style.display = 'block';
    };
    const hideComboList_ = idx => {
      const listEl = card.querySelector(`[data-combo-list="${idx}"]`);
      if (listEl) listEl.style.display = 'none';
    };
    const selectComboItem_ = (idx, value, name) => {
      const inputEl = card.querySelector(`input[data-combo="${idx}"]`);
      if (!inputEl) return;
      inputEl.value = name;
      inputEl.dataset.value = value;
      hideComboList_(idx);
      const chk = card.querySelector(`input[data-row="${idx}"]`);
      if (chk && !chk.checked) { chk.checked = true; refreshCount(); }
    };

    card.querySelector('#sp-confirm-tbody').addEventListener('input', e => {
      if (!e.target.matches('input[data-combo]')) return;
      const idx = +e.target.dataset.combo;
      // Typed text invalidates the previously selected value until the agent
      // clicks a suggestion — unless what they typed exactly matches an
      // option's name, which auto-resolves it without needing a click.
      const exact = (comboOpts[idx] || []).find(o => o.name.toLowerCase() === e.target.value.trim().toLowerCase());
      e.target.dataset.value = exact ? exact.value : '';
      renderComboList_(idx, e.target.value);
    });
    card.querySelector('#sp-confirm-tbody').addEventListener('focusin', e => {
      if (e.target.matches('input[data-combo]')) renderComboList_(+e.target.dataset.combo, e.target.value);
    });
    card.querySelector('#sp-confirm-tbody').addEventListener('focusout', e => {
      if (e.target.matches('input[data-combo]')) {
        const idx = +e.target.dataset.combo;
        setTimeout(() => hideComboList_(idx), 150); // let a pending mousedown-select land first
      }
    });
    card.querySelector('#sp-confirm-tbody').addEventListener('keydown', e => {
      if (!e.target.matches('input[data-combo]')) return;
      const idx = +e.target.dataset.combo;
      if (e.key === 'Escape') { hideComboList_(idx); e.target.blur(); }
      else if (e.key === 'Enter') {
        e.preventDefault();
        const first = filterCombo_(idx, e.target.value)[0];
        if (first) selectComboItem_(idx, first.value, first.name);
      }
    });
    // mousedown (not click) fires before the input's blur/focusout, so the
    // selection registers before the list gets hidden.
    card.querySelector('#sp-confirm-tbody').addEventListener('mousedown', e => {
      const item = e.target.closest('[data-combo-item]');
      if (!item) return;
      e.preventDefault();
      selectComboItem_(+item.dataset.comboItem, item.dataset.val, item.textContent);
    });

    card.querySelector('#sp-confirm-cancel').addEventListener('click', () => {
      overlay.remove(); onCancel?.();
    });
    card.querySelector('#sp-confirm-ok').addEventListener('click', () => {
      const selSet = new Set(rowChks().filter(c => c.checked).map(c => +c.dataset.row));
      // Sync live widget values back into each row before confirming — this is
      // what makes the After column actually editable (searchable dropdown for
      // tagger fields, text input for free-text fields), not just a static preview.
      card.querySelectorAll('input[data-combo]').forEach(comboEl => {
        const idx = +comboEl.dataset.combo;
        const row = rows[idx];
        if (!row) return;
        const val = comboEl.dataset.value || '';
        if (!val) { row.api = null; row.after = ''; return; }
        // row.api starts null for fields GCX Reply couldn't auto-resolve — build
        // it here (using the row's zdId) so a value the agent picks manually
        // actually gets submitted, not silently dropped.
        if (row.api) row.api.value = val;
        else if (row.zdId) row.api = { id: row.zdId, value: val };
        row.after = comboEl.value;
      });
      card.querySelectorAll('input[data-textinput]').forEach(inputEl => {
        const idx = +inputEl.dataset.textinput;
        const row = rows[idx];
        if (!row) return;
        row.after = inputEl.value;
        if (row.dom) {
          row.dom.after = inputEl.value;
          // Purchase Date's After input is now a native type="date" field
          // whose .value is always yyyy-mm-dd — the exact format apiVal
          // needs — so it's safe (and necessary, so a corrected date actually
          // reaches the PUT) to sync apiVal here too. Every other isDate-less
          // field has no separate apiVal concept, so this only affects Purchase Date.
          if (row.dom.isDate) row.dom.apiVal = inputEl.value;
        }
        if (row.api) row.api.value = inputEl.value;
      });
      overlay.remove();
      onConfirm(rows.filter((_, i) => selSet.has(i)));
    });
    overlay.addEventListener('click', e => { if (e.target === overlay) { overlay.remove(); onCancel?.(); } });
  }

  // ── Auto-fill status helpers ─────────────────────────────────────────────

  function setFillStatus(panel, msg) {
    const el = panel?.querySelector('#sp-fill-status');
    if (!el) return;
    el.textContent = msg;
    el.style.display = msg ? 'inline' : 'none';
  }

  function maybeShowAutoFill(panel) {
    const bar = panel?.querySelector('#sp-autofill-bar');
    const btn = panel?.querySelector('#sp-autofill-btn');
    if (bar && lastOrderData) {
      bar.style.display = 'block';
      // Only touch button state when not mid-fill (Preparing… / Filling…)
      if (btn && btn.textContent !== 'Preparing…' && btn.textContent !== 'Filling…') {
        btn.disabled = !_productReady;
        btn.textContent = _productReady ? 'Auto-Fill Form' : 'Loading…';
        btn.title = _productReady ? '' : 'Product info still loading — please wait';
      }
    }
    const mcfBar = panel?.querySelector('#sp-mcf-bar');
    if (mcfBar && lastOrderData) mcfBar.style.display = 'block';
  }

  // document.body.innerText forces a synchronous full-page reflow — cheap
  // once, but 3 separate call sites (autoDetectAll's initial scan, its own
  // fetchOrder ASIN fallback, autoFillTicket's device-candidate text) can
  // fire within the same ticket-load burst. Short TTL (not "cache until
  // navigation") deliberately: Zendesk tickets can get new content live
  // (agent replies, async comment load) while the agent stays on the page,
  // so a long-lived cache risks scanning stale text. 3s only helps the
  // near-simultaneous case; by the time Auto-Fill is clicked (seconds to
  // minutes later) the cache has always expired, so that path is unaffected.
  let _cachedBodyText = null, _cachedBodyTextAt = 0;
  function getCachedBodyText_() {
    const now = Date.now();
    if (!_cachedBodyText || now - _cachedBodyTextAt > 3000) {
      _cachedBodyText = document.body.innerText || '';
      _cachedBodyTextAt = now;
    }
    return _cachedBodyText;
  }

  // ── MCF: 티켓 본문 읽기 ────────────────────────────────────────────────────
  function getTicketBodyText_() {
    const m = location.pathname.match(/\/tickets\/(\d+)/);
    const pane = (m && document.querySelector(`[data-test-id="ticket-${m[1]}-standard-layout"]`)) || document.body;
    const inputText = [...pane.querySelectorAll('input, textarea')].map(el => el.value || '').join('\n');
    return (pane.innerText || '') + '\n' + inputText;
  }

  // ── AI 인입사유용 텍스트 추출 (고객 메시지 + Spigen CS 요약) ─────────────────
  function getAiInputText_() {
    const m = location.pathname.match(/\/tickets\/(\d+)/);
    const ticketId = m?.[1];
    const pane = (ticketId && document.querySelector(`[data-test-id="ticket-${ticketId}-standard-layout"]`)) || document.body;

    // Priority 1: Zendesk comment body elements (avoids sidebar metadata)
    const commentEls = pane.querySelectorAll(
      '[data-test-id*="comment-body"], [data-test-id*="comment-viewer"], ' +
      '[data-test-id*="event-message"], [data-test-id*="rich-text"], ' +
      '.rich_text, [data-garden-id*="comment"], [class*="CommentContent"]'
    );
    if (commentEls.length > 0) {
      const texts = [...commentEls].map(el => el.innerText.trim()).filter(t => t.length > 20);
      if (texts.length > 0) {
        return texts.slice(0, 4).join('\n\n').slice(0, 2500);
      }
    }

    // Priority 2: Find CS summary then extract surrounding conversation
    const full = (pane.innerText || '').replace(/[ \t]+/g, ' ').trim();
    if (!full) return '';

    const CS_NAME = 'Spigen Customer Service';
    const csIdx   = full.indexOf(CS_NAME);

    if (csIdx !== -1) {
      // Take content immediately BEFORE CS comment (actual conversation, not sidebar start)
      const start       = Math.max(0, csIdx - 1200);
      const customerPart = full.slice(start, csIdx).trim();
      const csPart       = full.slice(csIdx + CS_NAME.length, csIdx + CS_NAME.length + 600).trim();
      return [customerPart, csPart ? `[CS요약]\n${csPart}` : ''].filter(Boolean).join('\n\n---\n\n').slice(0, 2500);
    }

    // Priority 3: Skip first ~600 chars (usually sidebar metadata) and take the rest
    return full.slice(600, 2600);
  }

  // ── MCF: 티켓 본문에서 고객 주소 파싱 (MCF Autofill parseClipboard와 동일 로직) ─
  const _MCF_PHONE_CC = {
    PT:'+351',ES:'+34',DE:'+49',FR:'+33',IT:'+39',NL:'+31',SE:'+46',FI:'+358',
    BE:'+32',AT:'+43',IE:'+353',PL:'+48',RO:'+40',HU:'+36',GR:'+30',CZ:'+420',
    SK:'+421',LT:'+370',LV:'+371',EE:'+372',MT:'+356',CY:'+357',SI:'+386',
    HR:'+385',BG:'+359',LU:'+352',DK:'+45',GB:'+44',US:'+1',CA:'+1',IN:'+91',JP:'+81',
  };
  function _mcfNormCountry(tok) {
    if (!tok) return '';
    const t = tok.trim();
    if (/^UK$/i.test(t) || /^United\s*Kingdom$/i.test(t) || /^Grande.Bretagne$/i.test(t)) return 'GB';
    if (/^DEU?$/i.test(t) || /^Germany$/i.test(t) || /^Deutschland$/i.test(t) || /^Allemagne$/i.test(t)) return 'DE';
    if (/^Espa/i.test(t) || /^Spain$/i.test(t)) return 'ES';
    if (/^Portugal/i.test(t)) return 'PT';
    if (/^France$/i.test(t) || /^Francia$/i.test(t)) return 'FR';
    if (/^Italy$/i.test(t) || /^Itali[ae]$/i.test(t) || /^Italie$/i.test(t)) return 'IT';
    if (/^Netherlands$/i.test(t) || /^Holland$/i.test(t) || /^Pays.Bas$/i.test(t)) return 'NL';
    if (/^Belgium$/i.test(t) || /^Belgique$/i.test(t) || /^Belgien$/i.test(t)) return 'BE';
    if (/^Sweden$/i.test(t) || /^Sverige$/i.test(t)) return 'SE';
    if (/^Poland$/i.test(t) || /^Polen$/i.test(t)) return 'PL';
    if (/^Austria$/i.test(t) || /^[OÖ]sterreich$/i.test(t)) return 'AT';
    if (/^Ireland$/i.test(t) || /^Irland$/i.test(t)) return 'IE';
    if (/^Denmark$/i.test(t) || /^D[äa]nemark$/i.test(t)) return 'DK';
    if (/^Finland$/i.test(t) || /^Finnland$/i.test(t)) return 'FI';
    if (/^Greece$/i.test(t) || /^Griechenland$/i.test(t)) return 'GR';
    if (/^Romania$/i.test(t) || /^Roumanie$/i.test(t)) return 'RO';
    if (/^Hungary$/i.test(t) || /^Ungarn$/i.test(t)) return 'HU';
    if (/^Czech/i.test(t) || /^Tschechien$/i.test(t)) return 'CZ';
    if (/^Slovenia$/i.test(t) || /^Slowenien$/i.test(t)) return 'SI';
    if (/^Slovakia$/i.test(t) || /^Slowakei$/i.test(t)) return 'SK';
    if (/^Croatia$/i.test(t) || /^Kroatien$/i.test(t)) return 'HR';
    if (/^Bulgaria$/i.test(t) || /^Bulgarie$/i.test(t)) return 'BG';
    if (/^Estonia$/i.test(t) || /^Estland$/i.test(t)) return 'EE';
    if (/^Latvia$/i.test(t) || /^Lettland$/i.test(t)) return 'LV';
    if (/^Lithuania$/i.test(t) || /^Litauen$/i.test(t)) return 'LT';
    if (/^Malta$/i.test(t)) return 'MT';
    if (/^Cyprus$/i.test(t) || /^Zypern$/i.test(t)) return 'CY';
    if (/^Luxembourg$/i.test(t)) return 'LU';
    if (/^India$/i.test(t) || /^Inde$/i.test(t)) return 'IN';
    if (/^Japan$/i.test(t) || /^日本$/.test(t)) return 'JP';
    if (/^United\s*States$/i.test(t) || /^USA?$/i.test(t)) return 'US';
    if (/^Canada$/i.test(t)) return 'CA';
    return /^[A-Za-z]{2}$/.test(t) ? t.toUpperCase() : '';
  }
  // Maps a numbered-list line's label text to an address field, by keyword —
  // used so parseTicketAddress_ isn't fooled by renumbered/missing fields.
  function mcfFieldFromLabel_(label) {
    const l = (label || '').toLowerCase();
    if (/full\s*name|\bname\b/.test(l))              return 'name';
    if (/street|address\s*line|\baddress\b/.test(l)) return 'street';
    if (/\bcity\b|\btown\b/.test(l))                 return 'city';
    if (/\bstate\b|\bprovince\b|\bregion\b/.test(l))  return 'state';
    if (/zip|postal|post\s*code/.test(l))            return 'postal';
    if (/phone|tel(?:ephone)?\b|mobile/.test(l))     return 'phone';
    return null;
  }

  function parseTicketAddress_(txt) {
    if (!txt) return {};
    const t = txt.replace(/\r/g,'').replace(/[–—]/g,'-').replace(/ /g,' ').replace(/[ \t]+/g,' ').trim();
    const emailAll = t.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || [];
    const email = [...emailAll].reverse().find(e =>
      !/spigen\.com|zendesk\./i.test(e) &&
      !(/amazon\.(com|co\.uk|de|fr|es|it|nl|se)/i.test(e) && !/marketplace\.amazon\./i.test(e))
    ) || '';
    const asin = (t.match(/\bASIN\b[^\w]{0,5}(B[A-Z0-9]{9})\b/i) || [])[1] || '';
    const sku  = (t.match(/\bSKU\b[^\w]{0,5}([\w.-]{5,})/i) || [])[1] || '';
    let cRaw = (t.match(/Country\*?\s*[:\-：]\s*([^\n]+)/i) || [])[1]
            || (t.match(/^Country\*\n([A-Za-z]{2})\s*$/m) || [])[1]
            || (t.match(/국가\s*[:\-：]\s*([^\n]+)/i) || [])[1] || '';
    cRaw = cRaw.trim().replace(/\W+$/, '').trim();
    const blocks = [];
    let cur = null;
    const pushCur = () => { if (cur && Object.values(cur).some(Boolean)) blocks.push(cur); cur = null; };
    for (const line of t.split('\n').map(s => s.trim()).filter(Boolean)) {
      const m = line.match(/^\s*\(?(\d+)\)?[.)]\s*(.+)$/i);
      if (!m) continue;
      const idx = parseInt(m[1], 10);
      const rawContent = m[2].trim();

      // Identify the field by its label text (name/street/city/state/postal/
      // phone keywords) rather than trusting the number's position — if a
      // field is omitted and the list renumbers (e.g. City skipped, so State
      // becomes "3)" instead of "4)"), positional mapping alone would put
      // State's value in the city slot, postal in the state slot, etc. Falls
      // back to the original positional mapping only when the label doesn't
      // match any known keyword (e.g. no clear label at all).
      const colonM = rawContent.match(/^(.+?):\s*(.*)$/);
      const dashM  = !colonM && rawContent.match(/^(.+?)\s+--?\s*(.*)$/);
      const labelText = (colonM ? colonM[1] : dashM ? dashM[1] : '').trim();
      const val       = (colonM ? colonM[2] : dashM ? dashM[2] : rawContent).trim();
      const field = mcfFieldFromLabel_(labelText);

      // Only open a NEW block when line 1 is explicitly recognized as a
      // "Full Name" line (via mcfFieldFromLabel_) — gating purely on idx===1
      // let unrelated numbered lists elsewhere in the ticket (e.g. an
      // internal "1. 2. 3." processing/escalation note with no field labels)
      // masquerade as the start of an address block. Since the parser keeps
      // only the LAST completed block, such a trailing bogus block would
      // silently override a real, already-parsed address (confirmed live on
      // ticket #1000152805: a "Tier 2 처리 요청 사항" note after the real
      // address block clobbered it with Korean sentence fragments). Once a
      // block is open, unrelated numbered lists still can't be mistaken for
      // a continuation of it, since idx===1 is the only way in.
      if (idx === 1 && field === 'name') { pushCur(); cur = { name:'',street:'',city:'',state:'',postal:'',phone:'' }; }
      if (!cur) continue;
      if (field) {
        cur[field] = val;
      } else if (idx === 2) cur.street = val;
      else if   (idx === 3) cur.city   = val;
      else if   (idx === 4) cur.state  = val;
      else if   (idx === 5) cur.postal = val;
      else if   (idx === 6) cur.phone  = val;
      if (field === 'phone' || idx === 6) pushCur();
    }
    pushCur();
    const addr = blocks.length ? blocks[blocks.length - 1] : {};
    const phoneCountry = Object.entries(_MCF_PHONE_CC).find(([, cc]) => (addr.phone || '').includes(cc))?.[0] || '';
    const country = _mcfNormCountry(cRaw) || phoneCountry;
    return {
      name:   addr.name   || '', street: addr.street || '', city:   addr.city   || '',
      state:  addr.state  || '', postal: addr.postal || '', phone:  addr.phone  || '',
      email, q: asin || sku, country,
    };
  }

  // ── MCF: 주문 API + 티켓 본문 주소를 합쳐 해시 페이로드 생성 ─────────────────
  function buildMcfPayload_(panelEl, emailOverride, chosenAsin) {
    const o  = lastOrderData?.order   || {};
    const ad = lastOrderData?.address || {};
    const b  = lastOrderData?.buyer   || {};
    const itemAsins = (lastOrderData?.items || []).map(i => i.ASIN).filter(Boolean);
    const asin    = chosenAsin || itemAsins[0] || panelEl?.querySelector('#sp-asin-input')?.value.trim() || '';
    const orderId = panelEl?.querySelector('#sp-order-input')?.value.trim() || '';
    // 고객이 티켓에 직접 쓴 주소가 주문 API 주소보다 우선 (MCF 배송지이므로)
    const ta = parseTicketAddress_(getTicketBodyText_());
    const country = ta.country || ad.CountryCode || '';
    // Per-unit item price for MCF delivery cost ratio warning
    const rawAmt = parseFloat(o.OrderTotal?.Amount || 0);
    const totalQty = (lastOrderData?.items || []).reduce((s, i) => s + (i.QuantityOrdered || 1), 0);
    const itemPrice = rawAmt > 0 ? (rawAmt / Math.max(1, totalQty)).toFixed(2) : null;
    return {
      name:      ta.name   || b.BuyerName || o.BuyerInfo?.BuyerName || ad.Name || getTicketRequesterNameFromDom_() || '',
      street:    ta.street || ad.AddressLine1 || '',
      city:      ta.city   || ad.City || '',
      state:     ta.state  || ad.StateOrRegion || '',
      postal:    ta.postal || ad.PostalCode || '',
      phone:     ta.phone  || ad.Phone || '',
      email:     emailOverride || ta.email  || b.BuyerEmail || '',
      country,
      asin:      chosenAsin || ta.q || asin,
      orderId,
      region:    country === 'JP' ? 'JP' : 'global',
      itemPrice,
    };
  }

  // SP-API often lacks BuyerName PII permission (order loads fine, buyer name
  // comes back empty). Fallback: scrape the requester name straight off the
  // ticket page's own contact-card header (works regardless of SP-API perms).
  function getTicketRequesterNameFromDom_() {
    const el = document.querySelector(
      '[data-test-id="card-header-row-test-id"] [data-test-id="card-header-label-test-id"]'
    );
    return el ? el.textContent.trim() : '';
  }

  // The requester's contact card (read by getTicketRequesterNameFromDom_) is
  // lazy-rendered — it doesn't exist in the DOM until the "Customer context"
  // omnipanel tab has actually been opened at least once. Same pattern as
  // getCustomerContextEmail_ below, factored out so the name lookup doesn't
  // need to duplicate the open-panel logic (or the unrelated email parsing).
  // Returns true if this call had to switch tabs (i.e. Customer context
  // wasn't already open) — callers use that to know whether to switch back
  // afterward, since GCX Reply's own panel lives under the Apps tab and
  // switching away from it mid-Auto-Fill would hide it from the agent.
  async function ensureCustomerContextPanelOpen_() {
    // Poll for the button itself, not just click-once-if-found — this runs
    // as early as initial panel setup (loadUi().notes persisted-on case),
    // which can race ahead of Zendesk's own omnipanel sidebar rendering. A
    // single immediate querySelector silently gave up with nothing to click,
    // and nothing downstream ever got a chance to retry.
    // Confirmed live: the *entire* omnipanel sidebar (data-test-id=
    // "omnipanel-selector-wrapper", not just this one button) is sometimes
    // absent from the DOM altogether for several seconds — an intermittent
    // Zendesk-side render delay (~1 in 5 tickets), not a fixed short race.
    // 2s wasn't enough; give it up to 10s before giving up for real.
    const btn = await pollUntil_(
      () => document.querySelector('[data-test-id="omnipanel-selector-item-customer-context"]'),
      { maxTries: 40 } // ~40 * 250ms = 10s max wait
    );
    if (!btn) return false;
    if (btn.getAttribute('aria-pressed') !== 'true') {
      btn.click();
      await sleep(700);
      return true;
    }
    return false;
  }

  // Customer context 탭에서 고객 이메일 추출.
  // 탭이 접혀 있으면 먼저 클릭해서 열고, 패널 텍스트에서 이메일을 뽑아낸다.
  async function getCustomerContextEmail_() {
    const EXCLUDE = /spigen\.com|zendesk\./i;
    function extractEmail_(el) {
      const text = (el && (el.textContent || el.innerText)) || '';
      const all = text.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || [];
      return all.find(e => !EXCLUDE.test(e)) || '';
    }

    const btn = document.querySelector('[data-test-id="omnipanel-selector-item-customer-context"]');
    if (!btn) return '';

    if (btn.getAttribute('aria-pressed') !== 'true') {
      btn.click();
      await sleep(700);
    }

    // Try specific panel selectors first
    for (const sel of [
      '[data-test-id="customer-context-pane"]',
      '[data-test-id="omnipanel-content"]',
      '[data-garden-id="omnipanel.body"]',
      '[data-cy-test-element="customer-context-pane"]',
    ]) {
      const e = extractEmail_(document.querySelector(sel));
      if (e) return e;
    }

    // Walk up: button → span → li → ul → parent (omnipanel container)
    const container = btn.closest('ul')?.parentElement || btn.closest('[data-garden-id]')?.parentElement;
    const fromContainer = extractEmail_(container);
    if (fromContainer) return fromContainer;

    // Broadest fallback: right sidebar
    return extractEmail_(
      document.querySelector('[data-test-id="sidebar"]') ||
      document.querySelector('aside') ||
      document.querySelector('[role="complementary"]')
    );
  }

  // Shows a "what's new" popup once per version, unless the user dismissed it
  // for this exact version (24h snooze, or permanently until the next update).
  function loadUpdateDismiss_() {
    try { return JSON.parse(localStorage.getItem('gcx_update_dismiss') || 'null'); } catch { return null; }
  }
  function saveUpdateDismiss_(patch) {
    try { localStorage.setItem('gcx_update_dismiss', JSON.stringify(patch)); } catch (_) {}
  }

  function showUpdatePopupIfNeeded_() {
    const notes = CHANGELOG_[CURRENT_VERSION];
    if (!notes || !notes.length) return;
    const dismiss = loadUpdateDismiss_();
    if (dismiss && dismiss.version === CURRENT_VERSION) {
      if (dismiss.mode === 'permanent') return;
      if (dismiss.mode === 'snooze' && Date.now() < dismiss.until) return;
    }

    const overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:2147483647;display:flex;align-items:center;justify-content:center;';
    const box = document.createElement('div');
    box.style.cssText = 'background:#fff;border-radius:12px;padding:22px 26px;min-width:300px;max-width:380px;box-shadow:0 4px 24px rgba(0,0,0,.25);font-family:-apple-system,BlinkMacSystemFont,sans-serif;';

    const title = document.createElement('div');
    title.style.cssText = 'font-size:15px;font-weight:700;color:#2f3941;margin-bottom:10px;';
    title.textContent = `GCX Reply v${CURRENT_VERSION} 업데이트`;
    box.appendChild(title);

    const list = document.createElement('ul');
    list.style.cssText = 'margin:0 0 16px;padding-left:18px;font-size:12.5px;color:#2f3941;line-height:1.6;';
    notes.forEach(n => {
      const li = document.createElement('li');
      li.textContent = n;
      list.appendChild(li);
    });
    box.appendChild(list);

    const mkCheckbox = (id, text) => {
      const label = document.createElement('label');
      label.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;font-size:12px;color:#68737d;margin-bottom:6px;';
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.id = id;
      label.appendChild(chk);
      label.appendChild(document.createTextNode(text));
      box.appendChild(label);
      return chk;
    };
    // Mutually exclusive in effect: permanent wins if both are checked.
    const snoozeChk    = mkCheckbox('sp-update-snooze',    '24시간 동안 다시 보지 않기');
    const permanentChk = mkCheckbox('sp-update-permanent', '다음 업데이트까지 다시 보지 않기');
    snoozeChk.onchange    = () => { if (snoozeChk.checked) permanentChk.checked = false; };
    permanentChk.onchange = () => { if (permanentChk.checked) snoozeChk.checked = false; };

    const closeBtn = document.createElement('button');
    closeBtn.style.cssText = 'display:block;width:100%;text-align:center;background:#5ba4cf;border:none;color:#fff;cursor:pointer;font-size:13px;font-weight:600;padding:9px;margin-top:10px;border-radius:6px;';
    closeBtn.textContent = '확인';
    closeBtn.onclick = () => {
      if (permanentChk.checked) {
        saveUpdateDismiss_({ version: CURRENT_VERSION, mode: 'permanent' });
      } else if (snoozeChk.checked) {
        saveUpdateDismiss_({ version: CURRENT_VERSION, mode: 'snooze', until: Date.now() + 24 * 3600 * 1000 });
      }
      overlay.remove();
    };
    box.appendChild(closeBtn);

    overlay.appendChild(box);
    document.body.appendChild(overlay);
  }

  // One-time-per-browser "누구세요?" dropdown — resolves the selected name, or
  // null if dismissed without picking (caller falls back gracefully; MCF_Tracking's
  // backend itself also defaults an empty person to '김지우', so a cancel here
  // never blocks the actual MCF send, just leaves 담당자 attribution unset for
  // this one send).
  function showMcfPersonPicker_() {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:2147483646;display:flex;align-items:center;justify-content:center;';
      const box = document.createElement('div');
      box.style.cssText = 'background:#fff;border-radius:10px;padding:20px 24px;min-width:260px;box-shadow:0 4px 24px rgba(0,0,0,.25);font-family:-apple-system,BlinkMacSystemFont,sans-serif;';
      const title = document.createElement('div');
      title.style.cssText = 'font-size:14px;font-weight:600;color:#2f3941;margin-bottom:4px;';
      title.textContent = '누구세요?';
      const sub = document.createElement('div');
      sub.style.cssText = 'font-size:11px;color:#68737d;margin-bottom:14px;line-height:1.4;';
      sub.textContent = 'MCF 발송 로그의 담당자로 기록됩니다. 이 브라우저에 저장되며, 나중에 ⚙ 설정에서 바꿀 수 있습니다.';
      const select = document.createElement('select');
      select.style.cssText = 'width:100%;padding:6px 8px;font-size:13px;border:1px solid #d8dcde;border-radius:6px;margin-bottom:14px;box-sizing:border-box;';
      select.innerHTML = '<option value="" disabled selected>선택하세요</option>' +
        MCF_PERSON_OPTIONS.map(p => `<option value="${p}">${p}</option>`).join('');
      const confirmBtn = document.createElement('button');
      confirmBtn.style.cssText = 'display:block;width:100%;text-align:center;background:#1f73b7;color:#fff;border:none;padding:8px 12px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;';
      confirmBtn.textContent = '확인';
      confirmBtn.onclick = () => {
        if (!select.value) return;
        overlay.remove();
        resolve(select.value);
      };
      const cancel = document.createElement('button');
      cancel.style.cssText = 'display:block;width:100%;text-align:center;background:transparent;border:none;color:#68737d;cursor:pointer;font-size:12px;padding:6px;margin-top:6px;';
      cancel.textContent = '나중에';
      cancel.onclick = () => { overlay.remove(); resolve(null); };
      box.append(title, sub, select, confirmBtn, cancel);
      overlay.appendChild(box);
      overlay.onclick = e => { if (e.target === overlay) { overlay.remove(); resolve(null); } };
      document.body.appendChild(overlay);
    });
  }

  function showMcfAsinPicker_(asins) {
    return new Promise(resolve => {
      if (!asins || asins.length <= 1) { resolve((asins && asins[0]) || null); return; }
      const overlay = document.createElement('div');
      overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:2147483646;display:flex;align-items:center;justify-content:center;';
      const box = document.createElement('div');
      box.style.cssText = 'background:#fff;border-radius:10px;padding:20px 24px;min-width:280px;box-shadow:0 4px 24px rgba(0,0,0,.25);font-family:-apple-system,BlinkMacSystemFont,sans-serif;';
      const title = document.createElement('div');
      title.style.cssText = 'font-size:14px;font-weight:600;color:#2f3941;margin-bottom:4px;';
      title.textContent = 'MCF SKU 검색 ASIN 선택';
      const sub = document.createElement('div');
      sub.style.cssText = 'font-size:11px;color:#68737d;margin-bottom:14px;';
      sub.textContent = '티켓에서 여러 ASIN이 감지됐습니다. MCF 페이지에서 어떤 ASIN으로 SKU를 검색할까요?';
      box.appendChild(title);
      box.appendChild(sub);
      asins.forEach((a, i) => {
        const btn = document.createElement('button');
        btn.style.cssText = 'display:block;width:100%;text-align:left;background:#f8f9f9;border:1px solid #d8dcde;padding:8px 12px;margin-bottom:8px;border-radius:6px;cursor:pointer;font-size:13px;color:#2f3941;font-family:monospace;transition:all .15s;';
        btn.textContent = `${i + 1}. ${a}`;
        btn.onmouseenter = () => { btn.style.background = '#e8f5fd'; btn.style.borderColor = '#5ba4cf'; };
        btn.onmouseleave = () => { btn.style.background = '#f8f9f9'; btn.style.borderColor = '#d8dcde'; };
        btn.onclick = () => { overlay.remove(); resolve(a); };
        box.appendChild(btn);
      });
      const cancel = document.createElement('button');
      cancel.style.cssText = 'display:block;width:100%;text-align:center;background:transparent;border:none;color:#68737d;cursor:pointer;font-size:12px;padding:6px;margin-top:2px;';
      cancel.textContent = '취소';
      cancel.onclick = () => { overlay.remove(); resolve(null); };
      box.appendChild(cancel);
      overlay.appendChild(box);
      overlay.onclick = e => { if (e.target === overlay) { overlay.remove(); resolve(null); } };
      document.body.appendChild(overlay);
    });
  }

  async function sendToMCF(panel) {
    if (!lastOrderData) return;
    const status = panel.querySelector('#sp-mcf-status');
    const showSt = t => { if (status) { status.textContent = t; status.style.display = 'block'; } };

    showSt('고객 이메일 확인 중…');
    const ctxEmail = await getCustomerContextEmail_();

    // Collect all unique ASINs: SP-API order items + ASIN-labeled in ticket body
    const itemAsins = (lastOrderData?.items || []).map(i => i.ASIN).filter(Boolean);
    const bodyText  = getTicketBodyText_();
    const bodyAsins = [...new Set([...bodyText.matchAll(/\bASIN\b[^\w]{0,5}(B[A-Z0-9]{9})\b/gi)].map(m => m[1]))];
    const allAsins  = [...new Set([...itemAsins, ...bodyAsins])];

    let chosenAsin = null;
    if (allAsins.length > 1) {
      showSt('ASIN 선택 중…');
      chosenAsin = await showMcfAsinPicker_(allAsins);
      if (chosenAsin === null) { if (status) status.style.display = 'none'; return; }
    } else {
      chosenAsin = allAsins[0] || null;
    }

    const payload = buildMcfPayload_(panel, ctxEmail, chosenAsin);
    const encoded = btoa(encodeURIComponent(JSON.stringify(payload)));
    const _mcfW = window.open(getMcfBase_(payload.country) + '#spigen_mcf=' + encoded, '_blank');
    if (_mcfW) { try { _mcfW.name = 'spigen_mcf:' + encoded; } catch(e) {} }
    showSt('✓ MCF 탭 열림 — 자동입력 대기중');
    setTimeout(() => { if (status) status.style.display = 'none'; }, 4000);
  }

  // ── No Response Needed — direct Seller Central Messaging API ──────────────
  // ChannelReply subscription is being cancelled; this talks straight to
  // Seller Central's own internal messaging app (the same "Resolve Case"
  // action — confirmed by the team to be equivalent to NRN), using the
  // agent's own SC session. No ChannelReply dependency at all.
  //
  // Reverse-engineered from SC's messaging/v2 app:
  // 1. List the small set of currently-actionable cases (filterList=ACTION_NEEDED
  //    and REPORTED_UNRESOLVED — the only two valid filter values; together they
  //    match the "Cases requiring attention" bucket in the Message Center UI).
  // 2. For each case, GET its orderContext and match the target Amazon Order ID
  //    — SC has no direct order→case search, so this brute-forces the (small) list.
  // 3. POST /messaging/api/global/cases/{caseId}/caseStatus?marketplaceId={mp}
  //    body "SOLVED", using a fresh anti-csrftoken-a2z (comes back as a response
  //    header on any GET to /messaging/api/...) and x-resource-version (a static
  //    hash baked into SC's own JS bundle — scraped once, cached, re-scraped on
  //    failure since it changes whenever Amazon redeploys the messaging app).

  // Keyed by `base` (e.g. "https://sellercentral.amazon.de") — resource-version
  // hashes come from each marketplace origin's own JS bundle and differ between
  // origins. A single unkeyed cache meant relaying to a second marketplace
  // right after a first one (e.g. UK after DE) reused the wrong hash on the
  // first attempt every time, wasting a full round trip before the existing
  // one-shot retry corrected it.
  const _scResourceVersionByBase_ = new Map();

  function gmGet_(url, headers) {
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'GET', url, headers: headers || {}, timeout: 15000,
        onload:    res => resolve(res),
        onerror:   () => resolve(null),
        ontimeout: () => resolve(null),
      });
    });
  }

  function gmHeader_(res, name) {
    if (!res || !res.responseHeaders) return null;
    const m = res.responseHeaders.match(new RegExp('^' + name + ':\\s*(.+)$', 'im'));
    return m ? m[1].trim() : null;
  }

  function scMessagingBase_() {
    return scBaseUrl_(scDomain_(lastOrderData?.order?.SalesChannel, lastOrderData?.address?.CountryCode));
  }

  async function fetchScResourceVersion_(base) {
    if (_scResourceVersionByBase_.has(base)) return _scResourceVersionByBase_.get(base);
    const page = await gmGet_(`${base}/messaging/inbox`);
    const src  = page && page.responseText && page.responseText.match(/messaging\/v2\/main\.[a-f0-9]+\.js/);
    if (!src) return null;
    const bundle = await gmGet_(`${base}/${src[0]}`);
    const hash   = bundle && bundle.responseText && bundle.responseText.match(/[0-9a-f]{40}/);
    const version = hash ? hash[0] : null;
    _scResourceVersionByBase_.set(base, version);
    return version;
  }

  async function findScCaseForOrder_(base, orderId) {
    const lists = await Promise.all([
      gmGet_(`${base}/messaging/api/global/cases?filterList=ACTION_NEEDED&sortDirection=ASCENDING`),
      gmGet_(`${base}/messaging/api/global/cases?filterList=REPORTED_UNRESOLVED&sortDirection=ASCENDING`),
    ]);
    let csrfToken = null;
    const cases = [];
    for (const res of lists) {
      if (!res || res.status !== 200) continue;
      csrfToken = csrfToken || gmHeader_(res, 'anti-csrftoken-a2z');
      try { cases.push(...(JSON.parse(res.responseText).cases || [])); } catch (_) {}
    }
    const contexts = await Promise.all(cases.map(c =>
      gmGet_(`${base}/messaging/api/global/cases/${c.caseId}/orderContext?marketplaceId=${c.marketplaceId}`)
        .then(res => ({ c, res }))
    ));
    for (const { c, res } of contexts) {
      if (!res || res.status !== 200) continue;
      csrfToken = csrfToken || gmHeader_(res, 'anti-csrftoken-a2z');
      try {
        const data = JSON.parse(res.responseText);
        const orderIds = (data.orderContextList || []).flatMap(oc =>
          (oc.orderContextFields || []).filter(f => f.type === 'ORDER_ID').map(f => f.textValue));
        if (orderIds.includes(orderId)) return { caseId: c.caseId, marketplaceId: c.marketplaceId, csrfToken };
      } catch (_) {}
    }
    return null;
  }

  function resolveScCase_(base, caseId, marketplaceId, csrfToken, resourceVersion) {
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST',
        url:    `${base}/messaging/api/global/cases/${caseId}/caseStatus?marketplaceId=${marketplaceId}`,
        headers: {
          'content-type':        'application/json',
          'anti-csrftoken-a2z':  csrfToken,
          'x-requested-with':    csrfToken,
          'x-resource-version':  resourceVersion,
        },
        data:    'SOLVED',
        timeout: 15000,
        onload:    res => resolve(!!res && res.status >= 200 && res.status < 300),
        onerror:   () => resolve(false),
        ontimeout: () => resolve(false),
      });
    });
  }

  function ticketJsonAsync_(ticketId) {
    return new Promise(resolve => fetchTicketJson_(ticketId, resolve));
  }

  // Resolves the Seller Central case (+ marketplace base) for the CURRENT
  // ticket, trying the same priority chain relayAbmReply_ already uses for
  // the ABM auto-relay: the buyer proxy address's embedded case ID first
  // (works even with no order/ASIN at all — e.g. pre-purchase ABM
  // questions), then Amazon Order ID, then ASIN. Previously NRN only ever
  // tried the Order-ID path (findScCaseForOrder_ alone) and required
  // lastOrderData.order.AmazonOrderId to exist just to pick a marketplace
  // domain — so the button stayed disabled ("Only available once an Amazon
  // order is loaded") on any ABM ticket whose order hadn't resolved via
  // SP-API, even though the case is directly resolvable from the ticket's
  // own reply-to address regardless. Reported live: agent feedback on
  // ticket #1000154091 asking why NRN only activates when an Order ID
  // exists, when it should work for every ABM-received ticket.
  async function resolveNrnCase_() {
    const ticketId = location.pathname.match(/\/tickets\/(\d+)/)?.[1] || null;
    const ticket = ticketId ? await ticketJsonAsync_(ticketId) : null;
    const fromAddress  = ticket?.via?.source?.from?.address || '';
    const directCaseId = abmCaseIdFromAddress_(fromAddress) || (ticketId ? await abmCaseIdFromAudits_(ticketId).catch(() => null) : null);
    const abmDomain    = abmDomainFromAddress_(fromAddress);

    const orderId = lastOrderData?.order?.AmazonOrderId || null;
    const asin    = (lastOrderData?.items || []).map(i => i.ASIN).filter(Boolean)[0] || null;
    const base    = abmDomain ? scBaseUrl_(abmDomain) : scMessagingBase_();
    if (!base) return { match: null, orderId, base: null };

    let match = directCaseId ? await resolveScCaseDirect_(base, directCaseId).catch(() => null) : null;
    if (!match && orderId) match = await findScCaseForOrder_(base, orderId).catch(() => null);
    if (!match && asin)    match = await findScCaseForAsin_(base, asin).catch(() => null);
    return { match, orderId, base };
  }

  // Re-checks button enabled/disabled state from the API. Exposed on window so
  // markNoResponseNeeded() and the order-load handler can trigger a refresh.
  async function refreshNrnState_() {
    const setBtn = (disabled, title) => {
      const btn = document.querySelector('#sp-order-panel #sp-nrn-btn');
      if (!btn) return;
      btn.disabled = disabled;
      btn.title    = title;
    };
    const startTicketId = location.pathname.match(/\/tickets\/(\d+)/)?.[1] || null;
    setBtn(true, 'Checking…');
    const { match } = await resolveNrnCase_();
    if ((location.pathname.match(/\/tickets\/(\d+)/)?.[1] || null) !== startTicketId) return; // navigated mid-flight
    setBtn(!match, match ? 'Mark as No Response Needed' : 'No open Seller Central case found for this ticket (or already resolved)');
  }
  window.__gcxRefreshNrnState = refreshNrnState_;

  async function markNoResponseNeeded(panel) {
    const status = panel.querySelector('#sp-nrn-status');
    const show = (msg, color) => {
      if (!status) return;
      status.textContent = msg;
      status.style.color = color || '#27ae60';
      status.style.display = 'block';
      setTimeout(() => { status.style.display = 'none'; }, 4000);
    };

    const btn = panel.querySelector('#sp-nrn-btn');
    if (btn) btn.disabled = true;
    show('Finding case…', '#888');

    const { match, orderId, base } = await resolveNrnCase_();
    if (!match || !base) {
      show('No open Seller Central case found for this ticket', '#888');
      logStep_(`NRN: no matching SC case found${orderId ? ` for order ${orderId}` : ''}`);
      refreshNrnState_();
      return;
    }

    show('Marking…', '#888');
    let resourceVersion = await fetchScResourceVersion_(base).catch(() => null);
    let ok = resourceVersion && await resolveScCase_(base, match.caseId, match.marketplaceId, match.csrfToken, resourceVersion);
    if (!ok) {
      // Amazon may have redeployed the messaging app (stale resource-version) — refresh once and retry.
      _scResourceVersionByBase_.delete(base);
      resourceVersion = await fetchScResourceVersion_(base).catch(() => null);
      ok = resourceVersion && await resolveScCase_(base, match.caseId, match.marketplaceId, match.csrfToken, resourceVersion);
    }
    if (ok) {
      show('✓ Marked as "No response needed"');
      logStep_(`NRN: resolved SC case ${match.caseId}${orderId ? ` for order ${orderId}` : ''}`);
    } else {
      show('✗ Failed to mark', '#c0392b');
      logStep_(`NRN: failed to resolve SC case ${match.caseId}${orderId ? ` for order ${orderId}` : ''}`);
    }
    refreshNrnState_();
  }

  // ── ABM auto-relay: mirror public replies on ABM tickets to Seller Central ──
  // Zendesk's own comment submission fires a GraphQL "UpdateTicketMutation" via
  // window.fetch (confirmed live via network capture on ticket #1000153416 — NOT
  // the classic REST PUT). We patch fetch to observe (never block) that call: once
  // Zendesk's own request succeeds, if it carried a public, non-empty comment on an
  // ABM ticket (recipient = amazon@spigenhelp.zendesk.com), relay the same text to
  // the matching Seller Central case via the same reverse-engineered messaging API
  // the NRN feature already uses (POST .../artifacts/message). Case resolution tries
  // Amazon Order ID first (reuses findScCaseForOrder_), falls back to ASIN match
  // since some ABM tickets (pre-purchase questions) have no order at all.
  //
  // Request body shape for the send was read directly out of SC's own
  // messaging/v2 JS bundle (the sendMessageParams object literal built right
  // before the real POST .../artifacts/message dispatch — see sendScMessage_
  // for the field-by-field source) rather than sent live to confirm it, per
  // memory: feedback_no_test_messages_to_real_abm_customers. One field
  // (deviceView) couldn't be pinned down from the bundle in the time spent
  // looking and is a reasoned guess. If any field is still wrong, the send
  // fails cleanly (no artifact created) and logs a clear error — nothing is
  // ever delivered to a buyer on a failed/malformed attempt. The very next
  // real agent reply on a real ABM ticket is both the first real test and,
  // if it works, a genuine legitimate send — never a placeholder/probe
  // message.
  // (`_abmRelayHookInstalled` is declared near the top of the IIFE, not here —
  // see the call site for why.)
  function installAbmRelayHook_() {
    if (_abmRelayHookInstalled) return;
    _abmRelayHookInstalled = true;
    // Must patch unsafeWindow, not window: Tampermonkey runs userscripts in an
    // isolated sandbox where `window` is a proxy distinct from the real page's
    // window. Patching window.fetch here only rewrote the sandboxed copy —
    // Zendesk's own GraphQL client calls the real page's window.fetch and never
    // saw the patch, so the relay silently never fired (confirmed via live
    // testing: window.fetch.toString() on the real page showed Zendesk's own
    // Sentry-wrapped fetch, with no trace of this function, even though the
    // script itself ran fine — GM_xmlhttpRequest and DOM access work from the
    // sandbox regardless, which is why the panel and everything else worked).
    const origFetch = unsafeWindow.fetch;
    unsafeWindow.fetch = function (...args) {
      const [url, opts] = args;
      const isGraphqlPost = typeof url === 'string' && url.includes('/api/graphql') && opts && opts.method === 'POST';
      if (!isGraphqlPost) return origFetch.apply(this, args);
      let parsed = null;
      try { parsed = JSON.parse(opts.body); } catch (_) {}
      // Zendesk sometimes coalesces several GraphQL operations fired close
      // together (e.g. applying a macro that touches fields/tags AND submits
      // the reply+solve in one action) into a single POST whose body is a
      // JSON ARRAY of operations, not one operation object — confirmed live
      // 2026-09-15: ticket #1000161857 (Dutch ABM reply, solved via macro)
      // got its real public reply submitted to Zendesk fine, but left ZERO
      // trace in ABM_Relay_Log — not even a 'pending' row — while plainer
      // solve+reply tickets relayed fine. The old code did
      // `parsed?.operationName === 'UpdateTicketMutation'` directly, which is
      // always undefined on an array body (arrays have no .operationName),
      // so detection silently no-oped. Now normalizes to a list of ops and
      // matches each by the actual unique SHAPE we care about — a
      // variables.ticket.comment with isPublic + non-empty body.value —
      // instead of a specific operationName string, so it also survives
      // Zendesk ever renaming/versioning the mutation itself.
      const ops = Array.isArray(parsed) ? parsed : (parsed ? [parsed] : []);
      let ticketId = null, htmlBody = null;
      for (const op of ops) {
        const c = op?.variables?.ticket?.comment;
        if (c && c.isPublic && (c.body?.value || '').trim()) {
          ticketId = op.variables.id;
          htmlBody = c.body.value;
          break;
        }
      }
      const promise = origFetch.apply(this, args);
      if (ticketId && htmlBody) {
        // Clone before reading — the body stream can only be consumed once, and
        // Zendesk's own caller still needs to read the original `promise`'s response.
        // GraphQL returns HTTP 200 even on an application-level failure (the mutation
        // response is a union type: UpdateTicketSuccess | CreateOrUpdateTicketFailure),
        // so res.ok alone isn't enough — must confirm the comment actually saved
        // before relaying anything to a real buyer. A batched request's response is
        // itself an array in the same order, so check every entry for a success.
        promise.then(res => {
          if (!res.ok) return;
          res.clone().json().then(json => {
            const results = Array.isArray(json) ? json : [json];
            const success = results.some(r => r?.data?.updateTicket?.__typename === 'UpdateTicketSuccess');
            if (success) {
              relayAbmReply_(ticketId, htmlBody).catch(e => logStep_(`ABM relay: unhandled error — ${e.message}`));
            } else {
              logStep_('ABM relay: skipped — ticket update did not succeed');
            }
          }).catch(() => {});
        }).catch(() => {});
      }
      return promise;
    };
  }

  function abmDomainFromAddress_(address) {
    const m = (address || '').match(/@marketplace\.(amazon\.[a-z.]+)$/i);
    if (!m) return null;
    const domain = m[1].toLowerCase();
    return EU_SC_REDIRECT.has(domain) ? 'amazon.de' : domain;
  }

  // The Amazon buyer proxy address embeds the real SC case ID directly, e.g.
  // "8xqqjc9xb7w12pd+52f6fe61-05ba-40db-bba1-1cb13689b72a@marketplace.amazon.de".
  // This is far more reliable than order/ASIN matching — it works even when the
  // case has no order or ASIN attached (pre-purchase questions) and regardless
  // of whether the case is currently in an ACTION_NEEDED/REPORTED_UNRESOLVED
  // state, since order/ASIN matching only searches those two filtered lists.
  //
  // CONFIRMED LIVE 2026-07-30 (ticket #1000155523 investigation, see
  // [[abm_ticket_merge]]): Amazon's ABM email NEVER actually puts the "+uuid"
  // address in the From: header — ticket.via.source.from.address is always
  // the bare buyer-proxy address. The uuid only ever appears in the message's
  // OTHER original recipients, which Zendesk only exposes via
  // /tickets/{id}/audits.json (first audit's via.source.from.original_recipients),
  // never on the plain ticket object. So this function alone has never
  // actually resolved a case ID in production — every "working" direct
  // resolution was silently falling through to the Order-ID/ASIN fallback in
  // relayAbmReply_/resolveNrnCase_ instead. Kept as the fast, free first try
  // in case Amazon ever changes the email format; abmCaseIdFromAudits_ below
  // is the fallback that actually works today.
  const CASE_ID_RE = /\+([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})@marketplace\./i;
  function abmCaseIdFromAddress_(address) {
    const m = (address || '').match(CASE_ID_RE);
    return m ? m[1] : null;
  }

  // Fallback for the above: fetches the ticket's first audit event and scans
  // its original_recipients for a uuid-bearing address. One extra API call —
  // only paid when the free from.address check misses (which, per the above,
  // is effectively always for real ABM tickets today).
  async function abmCaseIdFromAudits_(ticketId) {
    const res = await gmGet_(`${location.origin}/api/v2/tickets/${ticketId}/audits.json?sort_order=asc`, { Accept: 'application/json' });
    if (!res || res.status !== 200) return null;
    let data;
    try { data = JSON.parse(res.responseText); } catch (_) { return null; }
    const first = (data.audits || [])[0];
    const recipients = (first && first.via && first.via.source && first.via.source.from
      && first.via.source.from.original_recipients) || [];
    for (const addr of recipients) {
      const c = abmCaseIdFromAddress_(addr);
      if (c) return c;
    }
    return null;
  }

  async function resolveScCaseDirect_(base, caseId) {
    const res = await gmGet_(`${base}/messaging/api/global/cases/${caseId}`);
    if (!res || res.status !== 200) return null;
    let data;
    try { data = JSON.parse(res.responseText); } catch (_) { return null; }
    const marketplaceId = data?.partyCase?.marketplaceId;
    if (!marketplaceId) return null;
    return { caseId, marketplaceId, csrfToken: gmHeader_(res, 'anti-csrftoken-a2z') };
  }

  // Same orderContext endpoint findScCaseForOrder_ uses (in the opposite
  // direction — there we already know the orderId and are searching for its
  // case; here we already know the caseId and want whatever order(s) it's
  // attached to). Returns an array (a case can carry more than one order
  // reference) — empty when the case genuinely has no order attached.
  async function fetchCaseOrderIds_(base, caseId, marketplaceId) {
    const res = await gmGet_(`${base}/messaging/api/global/cases/${caseId}/orderContext?marketplaceId=${marketplaceId}`);
    if (!res || res.status !== 200) return [];
    try {
      const data = JSON.parse(res.responseText);
      return (data.orderContextList || []).flatMap(oc =>
        (oc.orderContextFields || []).filter(f => f.type === 'ORDER_ID').map(f => f.textValue)).filter(Boolean);
    } catch (_) { return []; }
  }

  // Auto-fills the Zendesk Order ID field from the SC case's own order
  // association, for ABM tickets whose message text never mentioned an order
  // number at all (e.g. "there's a bubble under my screen protector" — no
  // order ID anywhere in the customer's typed text or the rest of the raw
  // Amazon email). ABM_TicketMerge's server-side auto-fill can only extract
  // an Order ID that's literally present in the raw email text; a live SC
  // session (this browser has one) can instead read the case's own order
  // context directly, which Seller Central always has regardless of what the
  // buyer typed. Confirmed live via ticket #1000154785 (SC showed the order,
  // the ABM ticket's Order ID field stayed blank for 9+ hours until an agent
  // ran Auto-Fill manually) and #1000154868.
  //
  // Per-ticket-visit guard (_abmOrderIdFillTicketId) avoids re-checking the
  // same ticket on every 1s NRN-state poll while the agent stays on it;
  // resets on navigation to a different ticket, so a fresh visit always
  // re-checks (safe/idempotent — no-ops the instant the field has a value,
  // written by anyone: this fill, the server-side one, or an agent).
  let _abmOrderIdFillTicketId = null;
  async function autoFillAbmOrderId_() {
    const ticketId = location.pathname.match(/\/tickets\/(\d+)/)?.[1] || null;
    if (!ticketId || ticketId === _abmOrderIdFillTicketId) return;
    const ticket = await ticketJsonAsync_(ticketId);
    if (!ticket || ticket.recipient !== 'amazon@spigenhelp.zendesk.com') return;
    if ((location.pathname.match(/\/tickets\/(\d+)/)?.[1] || null) !== ticketId) return; // navigated mid-flight
    const existing = (ticket.custom_fields || []).find(f => f.id === ZD.ORDER_ID);
    if (existing && existing.value) { _abmOrderIdFillTicketId = ticketId; return; }

    const fromAddress  = ticket?.via?.source?.from?.address || '';
    const directCaseId = abmCaseIdFromAddress_(fromAddress) || await abmCaseIdFromAudits_(ticketId).catch(() => null);
    if (!directCaseId) { _abmOrderIdFillTicketId = ticketId; return; } // no case ID to resolve from at all
    const abmDomain = abmDomainFromAddress_(fromAddress);
    const base = abmDomain ? scBaseUrl_(abmDomain) : scMessagingBase_();
    if (!base) return; // no marketplace resolvable yet — leave unguarded, retry on next poll

    const match = await resolveScCaseDirect_(base, directCaseId).catch(() => null);
    if (!match) return; // case not resolvable yet (e.g. no SC session) — retry later
    if ((location.pathname.match(/\/tickets\/(\d+)/)?.[1] || null) !== ticketId) return;

    const orderIds = await fetchCaseOrderIds_(base, match.caseId, match.marketplaceId).catch(() => []);
    const orderId = orderIds[0] || null;
    _abmOrderIdFillTicketId = ticketId; // either way, this ticket's been checked — case's order association won't change
    if (!orderId) return; // case genuinely has no order attached — nothing to fill
    if ((location.pathname.match(/\/tickets\/(\d+)/)?.[1] || null) !== ticketId) return;

    putZdTicket(ticketId, [{ id: ZD.ORDER_ID, value: orderId }], null, null, null);
    logStep_(`ABM: auto-filled Order ID ${orderId} from SC case ${match.caseId}`);
  }

  // ── Detected-language auto-correct ──────────────────────────────────────
  // Zendesk's own "Detected language" widget (shown above the first comment;
  // data-test-id="conversation-translator-dropdown") is an ML guess that
  // frequently misfires on short/templated messages — confirmed live on
  // ticket #1000155742: a plain French "Pouvez-vous fournir une facture pour
  // ma commande...?" was detected as "Yoruba". It is NOT a Zendesk ticket
  // field GCX Reply can PUT — it's a client-only React/downshift combobox
  // (selecting an option fires zero network requests, confirmed via a live
  // network trace on that same ticket), so correcting it can't corrupt any
  // stored ticket data; worst case is just a wrong-but-recoverable UI value
  // an agent can reselect. GCX Reply already knows the buyer's marketplace
  // country far more reliably than a short-text ML guess — either the live
  // SP-API shipping address once the order loads, or the Country* custom
  // field ABM_TicketMerge fills from the marketplace domain even when the
  // ticket has no order at all — so once a country is known, force-select
  // the matching language.
  //
  // Menu items render asynchronously after the trigger is clicked (~800ms
  // observed live) and the downshift instance number in its DOM ids
  // (e.g. "downshift-76-input"/"downshift-76-menu") is not stable across
  // page loads — derived from the input's own id instead of hardcoded.
  let _langCorrectedTicketId = null;
  async function autoCorrectDetectedLanguage_() {
    const curTicketId = () => location.pathname.match(/\/tickets\/(\d+)/)?.[1] || null;
    const ticketId = curTicketId();
    if (!ticketId || ticketId === _langCorrectedTicketId) return;

    let countryValue = COUNTRY_MAP[lastOrderData?.address?.CountryCode] || null;
    let countryKnown = !!countryValue;
    if (!countryKnown) {
      const ticket = await ticketJsonAsync_(ticketId);
      if (curTicketId() !== ticketId) return; // navigated mid-flight
      const cf = (ticket?.custom_fields || []).find(f => f.id === ZD.COUNTRY);
      if (cf && cf.value) { countryValue = cf.value; countryKnown = true; }
    }
    if (!countryKnown) return; // country not resolvable yet — retry on next poll

    const targetLang = LANGUAGE_BY_COUNTRY_VALUE[countryValue] || null;
    if (!targetLang) { _langCorrectedTicketId = ticketId; return; } // known country we deliberately don't map (e.g. BE)

    // Call sites only re-invoke this function on order-load or ticket
    // navigation, not on a fixed cadence — so wait a few seconds here for the
    // comment pane (and its translator widget) to finish rendering, rather
    // than relying on a lucky next trigger.
    await pollUntil_(() => document.querySelector('[data-test-id="conversation-translator-dropdown"]') || null, { intervalMs: 300, maxTries: 12 });
    if (curTicketId() !== ticketId) return; // navigated mid-flight
    const trigger    = document.querySelector('[data-test-id="conversation-translator-dropdown"]');
    const selectedEl = document.querySelector('[data-test-id="conversation-translator-selected-language"]');
    if (!trigger || !selectedEl) return; // widget never rendered (e.g. no inbound comment yet) — retry on next poll

    _langCorrectedTicketId = ticketId; // country + widget both confirmed — won't re-check this ticket again
    const before = selectedEl.textContent.trim();
    if (before === targetLang) return; // already correct

    const click_ = el => {
      const r = el.getBoundingClientRect();
      const o = { bubbles: true, cancelable: true, view: window, clientX: r.x + 5, clientY: r.y + 5 };
      el.dispatchEvent(new MouseEvent('mousedown', o));
      el.dispatchEvent(new MouseEvent('mouseup', o));
      el.dispatchEvent(new MouseEvent('click', o));
    };

    click_(trigger);
    const input = trigger.querySelector('input[id^="downshift-"]');
    const menu  = input && document.getElementById(input.id.replace(/-input$/, '-menu'));
    if (!menu) return;

    const items = await pollUntil_(() => {
      const opts = [...menu.querySelectorAll('li[role="option"]')];
      return opts.length ? opts : null;
    }, { intervalMs: 200, maxTries: 15 });
    if (curTicketId() !== ticketId) return; // navigated mid-flight

    const target = (items || []).find(li => li.textContent.trim() === targetLang);
    if (target) {
      click_(target);
      logStep_(`Detected language corrected: "${before}" → "${targetLang}" (${countryValue.toUpperCase()} marketplace)`);
    } else {
      click_(trigger); // close the menu without guessing further
    }
  }
  window.__gcxAutoCorrectLanguage = autoCorrectDetectedLanguage_;

  function htmlToPlainText_(html) {
    const div = document.createElement('div');
    div.innerHTML = (html || '').replace(/<br\s*\/?>/gi, '\n').replace(/<\/p>/gi, '\n');
    return (div.textContent || '').replace(/ /g, ' ').trim();
  }

  async function findScCaseForAsin_(base, asin) {
    const lists = await Promise.all([
      gmGet_(`${base}/messaging/api/global/cases?filterList=ACTION_NEEDED&sortDirection=ASCENDING`),
      gmGet_(`${base}/messaging/api/global/cases?filterList=REPORTED_UNRESOLVED&sortDirection=ASCENDING`),
    ]);
    let csrfToken = null;
    const cases = [];
    for (const res of lists) {
      if (!res || res.status !== 200) continue;
      csrfToken = csrfToken || gmHeader_(res, 'anti-csrftoken-a2z');
      try { cases.push(...(JSON.parse(res.responseText).cases || [])); } catch (_) {}
    }
    const match = cases.find(c => (c.asinIds || []).includes(asin));
    return match ? { caseId: match.caseId, marketplaceId: match.marketplaceId, csrfToken } : null;
  }

  // Body shape confirmed by reading SC's own messaging/v2 JS bundle directly
  // (the sendMessageParams object literal built right before the real dispatch
  // to POST .../artifacts/message, JSON.stringify'd as-is with no wrapper) —
  // not inferred from the read-model this time. `topicId`/`templateId` are
  // genuinely nullable in the real client too (ternary → null when unset);
  // `deviceView` value itself wasn't recoverable from the bundle in the time
  // spent looking — 'DESKTOP' is a reasoned guess consistent with the sibling
  // entryPoint enum (SELLER_DESKTOP/SELLER_MOBILE) and is the one remaining
  // unconfirmed field.
  function sendScMessage_(base, caseId, marketplaceId, orderId, subject, csrfToken, resourceVersion, text, attachments) {
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST',
        url: `${base}/messaging/api/global/cases/${caseId}/artifacts/message`,
        headers: {
          'content-type':       'application/json',
          'anti-csrftoken-a2z': csrfToken,
          'x-requested-with':   csrfToken,
          'x-resource-version': resourceVersion,
        },
        data: JSON.stringify({
          topicId:                     null,
          rawSubject:                  subject || '',
          rawMessageBody:              text,
          marketplaceId,
          orderId:                     orderId || null,
          attachments:                 attachments || [],
          messageEligibilityType:      null,
          messageEligibilityTypeReason: null,
          templateId:                  null,
          spellcheckUsed:              false,
          entryPoint:                  'SELLER_DESKTOP',
          deviceView:                  'DESKTOP',
        }),
        timeout: 15000,
        // {ok, why}: `why` feeds the relay log's LastError so a genuine Amazon
        // rejection (HTTP 4xx) is distinguishable from a missing SC session —
        // previously every failure collapsed to one indistinguishable string.
        onload:    res => resolve({ ok: !!res && res.status >= 200 && res.status < 300, why: res ? `HTTP ${res.status}` : 'no response' }),
        onerror:   () => resolve({ ok: false, why: 'network error' }),
        ontimeout: () => resolve({ ok: false, why: 'timeout' }),
      });
    });
  }

  // Total attachment size limit SC's own composer enforces (10 MB), read from
  // its JS bundle (Kp=10, Wp=1048576 → Kp*Wp bytes) — enforced here too so we
  // fail with a clear log/toast instead of a raw upload rejection.
  const ABM_ATTACH_MAX_TOTAL_BYTES = 10 * 1048576;

  function downloadAttachmentBlob_(url) {
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'GET', url, responseType: 'blob', timeout: 30000,
        onload:    res => resolve(res && res.status === 200 ? res.response : null),
        onerror:   () => resolve(null),
        ontimeout: () => resolve(null),
      });
    });
  }

  // Upload shape (multipart fields: attachment/attachmentFileName/marketplaceId/
  // caseId) and response shape ({id, fileName} used to build the send payload's
  // attachments array) both read directly from SC's own messaging/v2 JS bundle —
  // the jI() saga that POSTs to /messaging/api/global/attachments.
  function uploadScAttachment_(base, caseId, marketplaceId, csrfToken, resourceVersion, blob, fileName) {
    return new Promise(resolve => {
      const form = new FormData();
      form.append('attachment', blob, fileName);
      form.append('attachmentFileName', fileName);
      form.append('marketplaceId', marketplaceId);
      form.append('caseId', caseId);
      GM_xmlhttpRequest({
        method: 'POST',
        url: `${base}/messaging/api/global/attachments`,
        headers: {
          'anti-csrftoken-a2z': csrfToken,
          'x-requested-with':   csrfToken,
          'x-resource-version': resourceVersion,
        },
        data: form,
        timeout: 30000,
        onload: res => {
          if (!res || res.status < 200 || res.status >= 300) { resolve(null); return; }
          try {
            const data = JSON.parse(res.responseText);
            resolve(data?.id ? { id: data.id, fileName: data.fileName || fileName } : null);
          } catch (_) { resolve(null); }
        },
        onerror:   () => resolve(null),
        ontimeout: () => resolve(null),
      });
    });
  }

  async function fetchLatestCommentAttachments_(ticketId) {
    const res = await gmGet_(`${location.origin}/api/v2/tickets/${ticketId}/comments.json?sort_order=desc`, { Accept: 'application/json' });
    if (!res || res.status !== 200) return [];
    try {
      const latest = JSON.parse(res.responseText).comments?.[0];
      return latest?.attachments || [];
    } catch (_) { return []; }
  }

  // Attachments on a SPECIFIC comment (used by the reconciliation resend, where
  // the reply being delivered may not be the ticket's latest comment).
  async function fetchCommentAttachmentsById_(ticketId, commentId) {
    const res = await gmGet_(`${location.origin}/api/v2/tickets/${ticketId}/comments.json?sort_order=asc`, { Accept: 'application/json' });
    if (!res || res.status !== 200) return [];
    try {
      const c = (JSON.parse(res.responseText).comments || []).find(x => String(x.id) === String(commentId));
      return c?.attachments || [];
    } catch (_) { return []; }
  }

  // Uploads a given list of Zendesk attachments to the matched SC case,
  // skipping (with a log line, never a hard failure) any that fail individually
  // or that would push the total over SC's 10 MB composer limit — the message
  // text still sends even if an attachment can't be relayed.
  async function relayAbmAttachmentList_(base, caseId, marketplaceId, csrfToken, resourceVersion, zdAttachments) {
    if (!zdAttachments || !zdAttachments.length) return [];
    let totalBytes = 0;
    const uploaded = [];
    for (const att of zdAttachments) {
      if (totalBytes + (att.size || 0) > ABM_ATTACH_MAX_TOTAL_BYTES) {
        logStep_(`ABM relay: skipped attachment "${att.file_name}" — would exceed SC's 10 MB total limit`);
        continue;
      }
      const blob = await downloadAttachmentBlob_(att.content_url);
      if (!blob) { logStep_(`ABM relay: failed to download attachment "${att.file_name}" from Zendesk`); continue; }
      const up = await uploadScAttachment_(base, caseId, marketplaceId, csrfToken, resourceVersion, blob, att.file_name);
      if (!up) { logStep_(`ABM relay: failed to upload attachment "${att.file_name}" to Seller Central`); continue; }
      totalBytes += att.size || 0;
      uploaded.push(up);
    }
    return uploaded;
  }

  // Original real-time relay path: relay the just-created (latest) comment's
  // attachments.
  async function relayAbmAttachments_(base, caseId, marketplaceId, csrfToken, resourceVersion, ticketId) {
    const zdAttachments = await fetchLatestCommentAttachments_(ticketId);
    return relayAbmAttachmentList_(base, caseId, marketplaceId, csrfToken, resourceVersion, zdAttachments);
  }

  // Panel-independent toast — must work even if the floating panel is minimized,
  // docked-but-not-mounted, or not yet built (a reply can happen before init()'s
  // setTimeout(800ms) fires). On by default (whether a reply reached the buyer
  // is consequential), but mutable via the ⚙ "ABM alerts" setting — a failed
  // relay is still visible through the header ABM button's pending count and
  // keeps auto-retrying via the sweep either way.
  // `force` bypasses the alerts-off setting — used for feedback on an action
  // the agent explicitly clicked (e.g. Retry in the ABM panel), which should
  // always answer them even when automatic alerts are muted.
  function abmToast_(msg, ok, force) {
    if (!force && !abmAlertsEnabled_()) return;
    const el = document.createElement('div');
    el.textContent = msg;
    el.style.cssText = `position:fixed;top:16px;right:16px;z-index:2147483647;
      background:${ok ? '#1e8e3e' : '#c0392b'};color:#fff;padding:10px 16px;
      border-radius:8px;font:600 13px -apple-system,sans-serif;max-width:360px;
      box-shadow:0 4px 16px rgba(0,0,0,.25);`;
    document.body.appendChild(el);
    setTimeout(() => el.remove(), ok ? 5000 : 9000);
  }

  async function relayAbmReply_(ticketId, htmlBody) {
    // Keyed on NORMALIZED plain text, not raw HTML — confirmed live 2026-09-16:
    // 56+ ticket pairs in ABM_Relay_Log show the identical visible message sent
    // twice to a real buyer (e.g. #1000161645, two relayKeys 14s apart), even
    // though this hash-based claim already existed. Checked Zendesk directly:
    // only ONE real comment ever existed for each pair — so two separate
    // relayAbmReply_ invocations fired for the same single reply, and the raw-
    // HTML hash let both through because the two invocations' htmlBody strings
    // evidently weren't byte-identical (Zendesk's composer can re-serialize the
    // same visible text with cosmetically different markup — e.g. a second
    // GraphQL call the hook also matches). Hashing the same normalized-text
    // form the GAS-side reconciliation dedup already uses successfully
    // (_normalizeAbmTextForDedup_ / normalizeForDedup_) closes this regardless
    // of which exact markup difference caused it.
    const plainText = htmlToPlainText_(htmlBody);
    const sendClaimKey = `${ticketId}_${_abmSimpleHash_(_normalizeAbmTextForDedup_(plainText))}`;
    if (!(await claimAbmSend_(sendClaimKey))) {
      logStep_(`ABM relay: skipped — this exact reply on ticket #${ticketId} was already relayed moments ago (duplicate-trigger guard)`);
      return;
    }
    const ticketRes = await gmGet_(`${location.origin}/api/v2/tickets/${ticketId}.json`, { Accept: 'application/json' });
    if (!ticketRes || ticketRes.status !== 200) { logStep_('ABM relay: could not load ticket'); return; }
    let ticket;
    try { ticket = JSON.parse(ticketRes.responseText).ticket; } catch (_) { logStep_('ABM relay: ticket parse error'); return; }
    if (ticket.recipient !== 'amazon@spigenhelp.zendesk.com') return; // not an ABM ticket — nothing to do

    const fromAddress = ticket.via?.source?.from?.address || '';
    const domain = abmDomainFromAddress_(fromAddress);
    if (!domain) {
      logStep_(`ABM relay: could not resolve marketplace domain from "${fromAddress}"`);
      abmToast_('⚠ ABM relay: could not resolve marketplace — reply manually in Seller Central', false);
      return;
    }
    const base = scBaseUrl_(domain);

    const desc    = ticket.description || '';
    const orderId = desc.match(/\b(\d{3}-\d{7}-\d{7})\b/)?.[1] || null;
    const asin    = desc.match(/\b(B[A-Z0-9]{9})\b/)?.[1] || null;
    const directCaseId = abmCaseIdFromAddress_(fromAddress) || await abmCaseIdFromAudits_(ticketId).catch(() => null);

    logStep_(`ABM relay: resolving SC case (direct=${directCaseId || '—'}, order=${orderId || '—'}, asin=${asin || '—'})…`);
    const resolveMatch = async () => {
      let m = directCaseId ? await resolveScCaseDirect_(base, directCaseId).catch(() => null) : null;
      if (!m && orderId) m = await findScCaseForOrder_(base, orderId).catch(() => null);
      if (!m && asin) m = await findScCaseForAsin_(base, asin).catch(() => null);
      return m;
    };
    let match = await resolveMatch();
    if (!match) {
      // A marketplace origin the browser hasn't hit recently in this session
      // (e.g. replying on a UK/IE ticket right after a run of DE/FR/IT ones)
      // can cold-fail the first request even with a perfectly valid login —
      // confirmed live: a case that failed here resolved fine on a plain
      // manual re-request seconds later. One retry after a short delay
      // clears this without ever bothering the agent for the transient case.
      logStep_('ABM relay: case resolution failed once — retrying after a short delay…');
      await sleep(1200);
      match = await resolveMatch();
    }
    if (!match) {
      logStep_('ABM relay: no matching SC case found — reply is only in Zendesk, reply must also be sent in Seller Central manually');
      abmToast_('⚠ ABM relay: no matching Seller Central case found — reply manually in Seller Central', false);
      return;
    }

    // Unique per relay invocation (not just per ticket) — a single ticket can
    // get multiple agent replies, each triggering its own separate relay. An
    // earlier version keyed the log by bare ticketId, so a second reply's log
    // entry silently overwrote the first reply's row — confirmed live: two
    // replies each on tickets #1000153623/#1000153609 left only one `success`
    // row per ticket, with no way to tell which send it actually corresponded
    // to. Threaded through every log call below so the pending→final update
    // for THIS message never collides with another message on the same ticket.
    const relayKey = `${ticketId}_${Date.now()}`;
    // Logged as "pending" before any send attempt — not just on final success/
    // failure — so a hard reload or closed tab mid-flight (the multi-attempt
    // send with deliberate retry delays can take several seconds) still
    // leaves a recoverable row for the sweep to pick up later, instead of
    // vanishing with zero record of the attempt ever happening.
    logAbmRelayResult_(relayKey, ticketId, match.caseId, domain, 'pending', 0, '', plainText);
    let resourceVersion = await fetchScResourceVersion_(base).catch(() => null);
    // Upload attachments once, before any send attempt, and reuse the same
    // uploaded refs across retries — re-uploading on every retry would
    // duplicate them in Seller Central for no benefit.
    const attachmentRefs = resourceVersion
      ? await relayAbmAttachments_(base, match.caseId, match.marketplaceId, match.csrfToken, resourceVersion, ticketId).catch(() => [])
      : [];
    const { ok, attempts, lastWhy } = await sendAbmMessageWithRetry_(base, match, orderId, ticket.subject, plainText, attachmentRefs);
    logStep_(ok
      ? `ABM relay: sent to SC case ${match.caseId} (attempt ${attempts})`
      : `ABM relay: FAILED to send to SC case ${match.caseId} after ${attempts} attempts (${lastWhy}) — will keep retrying automatically; reply is only in Zendesk for now, agent should also reply in Seller Central manually if urgent`);
    abmToast_(ok
      ? '✓ Also sent to the Amazon buyer via Seller Central'
      : '✗ Could NOT relay to Amazon after 3 attempts — will keep retrying automatically in the background; reply manually in Seller Central if urgent', ok);
    logAbmRelayResult_(relayKey, ticketId, match.caseId, domain, ok ? 'success' : 'failed', attempts, ok ? '' : `exhausted retries — ${lastWhy}`, plainText);
    if (!ok) updateAbmBadge_();
  }

  // A prior version only refreshed the resource-version hash on retry and kept
  // reusing the csrfToken captured back when the SC case was first resolved —
  // but by the time the send actually fires (case resolution + a possible
  // 1200ms resolve-retry + attachment upload/download round trips +
  // resourceVersion fetch), several seconds can have passed, and Amazon's
  // anti-csrftoken-a2z can go stale in that window. That reproduced exactly
  // what agents reported: a reply silently fails once, then succeeds on a
  // plain manual resend — because a fresh compose+submit re-resolves
  // everything from scratch, including a brand-new csrfToken, while the old
  // retry path kept reusing the same (possibly already-stale) one. Refreshing
  // both the resourceVersion AND the csrfToken on every retry (up to 3 total
  // attempts) closes that gap.
  async function sendAbmMessageWithRetry_(base, match, orderId, subject, plainText, attachmentRefs) {
    const MAX_ATTEMPTS = 3;
    let csrfToken = match.csrfToken;
    let lastWhy = '';
    for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
      const resourceVersion = await fetchScResourceVersion_(base).catch(() => null);
      if (!resourceVersion) {
        lastWhy = 'resource-version fetch failed';
      } else {
        const sent = await sendScMessage_(base, match.caseId, match.marketplaceId, orderId, subject, csrfToken, resourceVersion, plainText, attachmentRefs);
        if (sent.ok) return { ok: true, attempts: attempt };
        lastWhy = `send rejected (${sent.why})`;
      }
      if (attempt < MAX_ATTEMPTS) {
        _scResourceVersionByBase_.delete(base);
        await sleep(800);
        const refreshed = await resolveScCaseDirect_(base, match.caseId).catch(() => null);
        if (refreshed) csrfToken = refreshed.csrfToken;
      }
    }
    return { ok: false, attempts: MAX_ATTEMPTS, lastWhy };
  }

  function logAbmRelayResult_(relayKey, ticketId, caseId, marketplace, status, attempts, lastError, messageText, commentId) {
    GM_xmlhttpRequest({
      method: 'POST',
      url: GAS_URL,
      headers: { 'Content-Type': 'application/json' },
      data: JSON.stringify({ action: 'logAbmRelay', relayKey, ticketId, commentId: commentId || '', caseId, marketplace, status, attempts, lastError, messageText }),
      timeout: 15000,
      onload: () => {}, onerror: () => {}, ontimeout: () => {},
    });
  }

  // Persistent (not auto-dismissing) on-page indicator of any ABM replies that
  // haven't yet been confirmed delivered to the buyer — unlike abmToast_, this
  // survives ticket navigation and page reloads, since it's re-derived from
  // the durable GAS-backed log rather than in-memory state. Dismissible via ✕
  // (stays hidden until the SET of pending items changes — a brand-new failure
  // re-shows it; the same known backlog doesn't nag again). Clicking the text
  // opens the status panel (per-ticket detail + editable statuses) instead of
  // blind-firing a retry sweep.
  let _abmBadgeEl_ = null;
  let _abmDismissedSig_ = null;
  const _abmPendingSig_ = pending => pending.map(p => p.relayKey || p.ticketId).sort().join('|');
  // Agents reported the floating "N ABM reply not yet delivered" badge
  // popping up the instant they hit send — confusing, because a brand-new
  // 'pending' row is almost always just a completely normal relay still in
  // flight (network round trip + up to 3 retries with backoff), not a real
  // failure. The relay log's own Timestamp column isn't a reliable proxy for
  // "time since replied" either — upsertAbmRelayLog_ rewrites it on every
  // retry attempt, so a message stuck for hours but retried every 5 min would
  // never look stale by that measure. Only the real signal — the Zendesk
  // reply comment's own created_at — answers "how long has this actually been
  // undelivered" correctly. So the alert (badge + floating popup) only
  // surfaces a pending item once its ORIGINATING REPLY is over an hour old.
  // v3.6.4: the panel's own "Undelivered" list now applies this same 1hr
  // gate too (previously left unfiltered on purpose, showing brand-new
  // in-flight rows as "failed" and confusing agents into thinking a normal
  // send had failed) — per explicit user request, only genuinely stuck
  // (>1hr) rows count as Undelivered anywhere in the UI.
  const ABM_ALERT_MIN_AGE_MS = 60 * 60 * 1000; // 1 hour

  async function abmCommentCreatedAt_(ticketId, commentId) {
    if (!commentId) return null;
    const res = await gmGet_(`${location.origin}/api/v2/tickets/${ticketId}/comments.json?sort_order=desc`, { Accept: 'application/json' });
    if (!res || res.status !== 200) return null;
    try {
      const c = (JSON.parse(res.responseText).comments || []).find(x => String(x.id) === String(commentId));
      return c ? new Date(c.created_at).getTime() : null;
    } catch (_) { return null; }
  }

  async function abmFilterAlertWorthy_(pending) {
    const now = Date.now();
    const out = [];
    for (const p of pending) {
      // Fall back to the log row's own timestamp only for legacy rows with no
      // CommentId (pre-dates that column) — best available signal there.
      let ts = await abmCommentCreatedAt_(p.ticketId, p.commentId);
      if (!ts) ts = new Date(p.timestamp).getTime();
      if (ts && (now - ts) >= ABM_ALERT_MIN_AGE_MS) out.push(p);
    }
    return out;
  }

  // Static (never flashing/popping) pending count on the panel header's ABM
  // button — the always-available entry point to the relay-status panel, and
  // the only indicator agents see when they've muted ABM alerts in ⚙.
  function updateAbmHeaderCount_(pendingCount) {
    const btn = document.getElementById('sp-abm-log-btn');
    if (!btn) return;
    btn.textContent = pendingCount > 0 ? `ABM(${pendingCount})` : 'ABM';
    btn.classList.toggle('sp-abm-has-pending', pendingCount > 0);
    btn.title = pendingCount > 0
      ? `${pendingCount} ABM ${pendingCount === 1 ? 'reply' : 'replies'} not yet delivered — click for details`
      : 'ABM relay log';
  }
  async function updateAbmBadge_() {
    const res = await gmGet_(`${GAS_URL}?action=abmRelayPending&clientVersion=${CURRENT_VERSION}`);
    let pending = [];
    if (res && res.status === 200) {
      try { pending = JSON.parse(res.responseText).pending || []; } catch (_) {}
    }
    const alertWorthy = await abmFilterAlertWorthy_(pending);
    updateAbmHeaderCount_(alertWorthy.length);
    const sig = _abmPendingSig_(alertWorthy);
    if (!alertWorthy.length || sig === _abmDismissedSig_ || !abmAlertsEnabled_()) {
      if (_abmBadgeEl_) { _abmBadgeEl_.remove(); _abmBadgeEl_ = null; }
      return;
    }
    if (!_abmBadgeEl_) {
      _abmBadgeEl_ = document.createElement('div');
      _abmBadgeEl_.id = 'sp-abm-badge';
      _abmBadgeEl_.style.cssText = `position:fixed;top:64px;right:16px;z-index:2147483647;
        background:#c0392b;color:#fff;padding:8px 12px;border-radius:20px;
        font:600 12px -apple-system,sans-serif;box-shadow:0 4px 16px rgba(0,0,0,.3);
        display:flex;align-items:center;gap:10px;pointer-events:auto;`;
      const txt = document.createElement('span');
      txt.id = 'sp-abm-badge-txt';
      txt.style.cursor = 'pointer';
      txt.title = 'Click for details';
      txt.onclick = () => showAbmPanel_();
      const x = document.createElement('span');
      x.textContent = '✕';
      x.title = 'Dismiss (re-appears if a NEW failure occurs)';
      x.style.cssText = 'cursor:pointer;opacity:.75;padding:0 2px;font-size:13px;';
      x.onclick = ev => {
        ev.stopPropagation();
        _abmDismissedSig_ = _abmBadgeEl_.dataset.sig || null;
        _abmBadgeEl_.remove();
        _abmBadgeEl_ = null;
      };
      _abmBadgeEl_.append(txt, x);
      document.body.appendChild(_abmBadgeEl_);
    }
    _abmBadgeEl_.dataset.sig = sig;
    _abmBadgeEl_.querySelector('#sp-abm-badge-txt').textContent =
      `⚠ ${alertWorthy.length} ABM ${alertWorthy.length === 1 ? 'reply' : 'replies'} not yet delivered`;
  }

  // ── ABM relay status panel ────────────────────────────────────────────────
  // Opened from the badge: lists Undelivered and Delivered relays separately,
  // each row showing ticket link + time + message snippet, with editable
  // status — Retry / "Mark delivered" on undelivered rows (takes it out of the
  // sweep when the agent has confirmed or manually handled it in SC), and
  // "Mark undelivered" on delivered rows (re-queues one the agent knows never
  // actually arrived, e.g. the pre-v3.3.10 rows whose per-ticket log entries
  // were overwritten and can't be trusted).
  let _abmPanelEl_ = null;
  const _abmEsc_ = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function postAbmStatus_(relayKey, ticketId, status) {
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST',
        url: GAS_URL,
        headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify({ action: 'setAbmRelayStatus', relayKey, ticketId, status }),
        timeout: 15000,
        onload: () => resolve(true), onerror: () => resolve(false), ontimeout: () => resolve(false),
      });
    });
  }

  async function showAbmPanel_() {
    if (_abmPanelEl_) { _abmPanelEl_.remove(); _abmPanelEl_ = null; }
    const panel = document.createElement('div');
    _abmPanelEl_ = panel;
    panel.id = 'sp-abm-panel';
    panel.style.cssText = `position:fixed;top:64px;right:16px;z-index:2147483647;width:440px;
      max-height:75vh;overflow-y:auto;background:#fff;color:#333;border-radius:12px;
      box-shadow:0 8px 32px rgba(0,0,0,.35);font:13px -apple-system,sans-serif;`;
    panel.innerHTML = `
      <div style="position:sticky;top:0;background:#f6f7f9;padding:10px 14px;display:flex;
                  align-items:center;justify-content:space-between;border-bottom:1px solid #e3e5e8;">
        <b>ABM Relay Status</b>
        <span>
          <button id="sp-abm-retry-sel" disabled style="border:1px solid #1e8e3e;background:#e6f4ea;color:#1e8e3e;
            border-radius:6px;padding:3px 10px;cursor:pointer;font-size:12px;margin-right:6px;">Retry selected</button>
          <button id="sp-abm-retry-all" style="border:1px solid #ccc;background:#fff;border-radius:6px;
            padding:3px 10px;cursor:pointer;font-size:12px;margin-right:8px;">Retry all</button>
          <span id="sp-abm-panel-x" style="cursor:pointer;font-size:15px;padding:2px 4px;">✕</span>
        </span>
      </div>
      <div id="sp-abm-panel-body" style="padding:10px 14px;">Loading…</div>`;
    document.body.appendChild(panel);
    panel.querySelector('#sp-abm-panel-x').onclick = () => { panel.remove(); _abmPanelEl_ = null; };
    panel.querySelector('#sp-abm-retry-all').onclick = async ev => {
      ev.target.disabled = true; ev.target.textContent = 'Retrying…';
      await sweepAbmRelayFailures_();
      await renderAbmPanelBody_(panel);
      ev.target.disabled = false; ev.target.textContent = 'Retry all';
    };
    // Retry only the rows the agent checked — reads the CURRENT body's
    // checkboxes at click time (the body re-renders independently of this
    // sticky header; renderAbmPanelBody_ stashes the latest undelivered
    // array on panel._abmUndelivered for exactly this).
    panel.querySelector('#sp-abm-retry-sel').onclick = async ev => {
      const items = [...panel.querySelectorAll('#sp-abm-panel-body [data-abm-sel]:checked')]
        .map(cb => (panel._abmUndelivered || [])[Number(cb.dataset.abmSel)])
        .filter(Boolean);
      if (!items.length) return;
      ev.target.disabled = true; ev.target.textContent = 'Retrying…';
      let okCount = 0;
      for (const item of items) {
        if (await trySendPendingItem_(item)) okCount++;
      }
      abmToast_(okCount === items.length
        ? `✓ ${okCount}/${items.length} selected ${items.length === 1 ? 'reply' : 'replies'} delivered`
        : `${okCount}/${items.length} delivered — check the Seller Central login for the failed ${items.length - okCount === 1 ? 'one' : 'ones'}`,
        okCount === items.length, true);
      await renderAbmPanelBody_(panel);
      updateAbmBadge_();
      ev.target.textContent = 'Retry selected';
    };
    await renderAbmPanelBody_(panel);
  }

  async function renderAbmPanelBody_(panel) {
    const body = panel.querySelector('#sp-abm-panel-body');
    // Undelivered MUST come from a full-sheet scan (abmRelayPending — same
    // source updateAbmBadge_'s header count already uses), not a slice of
    // the account-wide log sorted by recency. abmRelayAll&limit=40 only
    // covers the last few hours of activity at current volume (470+ total
    // rows) — an older stuck/failed row falls out of that window entirely,
    // so this panel silently showed "None 🎉" while Seller Central still had
    // a real unsent message and the header badge (correctly, via the
    // unlimited abmRelayPending) disagreed. Confirmed live 2026-07-29.
    // Delivered history stays on the capped abmRelayAll — that's just a
    // recent-activity convenience list, not correctness-critical.
    const [pendingRes, allRes] = await Promise.all([
      gmGet_(`${GAS_URL}?action=abmRelayPending&clientVersion=${CURRENT_VERSION}`),
      gmGet_(`${GAS_URL}?action=abmRelayAll&limit=40`),
    ]);
    let undelivered = [];
    let allRows = [];
    if (pendingRes && pendingRes.status === 200) {
      try { undelivered = JSON.parse(pendingRes.responseText).pending || []; } catch (_) {}
    }
    if (allRes && allRes.status === 200) {
      try { allRows = JSON.parse(allRes.responseText).rows || []; } catch (_) {}
    }
    // Guard against junk/migration rows: a real row always has a numeric ticket id.
    undelivered = undelivered.filter(r => /^\d+$/.test(String(r.ticketId)));
    allRows = allRows.filter(r => /^\d+$/.test(String(r.ticketId)));
    // v3.6.4: only rows genuinely stuck for over an hour count as
    // "Undelivered" here — same abmFilterAlertWorthy_ gate the header
    // badge already uses (reads each row's real Zendesk reply created_at,
    // not the relay log's own Timestamp, which upsertAbmRelayLog_ rewrites
    // on every retry). A brand-new relay still mid-flight (normal network
    // round trip + up to 3 retries) is not a failure and shouldn't render
    // as one just because an agent opened this panel early.
    undelivered = await abmFilterAlertWorthy_(undelivered);
    const delivered = allRows.filter(r => r.status === 'success');
    const fmtTime = ts => { const d = new Date(ts); return isNaN(d) ? '' : d.toLocaleString('en-GB', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }); };
    const rowHtml = (r, i, delivered_) => `
      <div style="padding:8px 0;border-bottom:1px solid #eee;">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
          <span style="display:flex;align-items:center;gap:6px;">
            ${delivered_ ? '' : `<input type="checkbox" data-abm-sel="${i}" style="cursor:pointer;margin:0;" title="Select for Retry selected"/>`}
            <span>
              <a href="/agent/tickets/${_abmEsc_(r.ticketId)}" target="_blank"
                 style="color:#1f73b7;font-weight:600;text-decoration:none;">#${_abmEsc_(r.ticketId)}</a>
              <span style="color:#888;margin-left:6px;">${fmtTime(r.timestamp)}</span>
              <span style="margin-left:6px;padding:1px 7px;border-radius:9px;font-size:11px;font-weight:600;
                background:${delivered_ ? '#e6f4ea' : '#fdecea'};color:${delivered_ ? '#1e8e3e' : '#c0392b'};">
                ${delivered_ ? 'delivered' : _abmEsc_(r.status || 'failed')}</span>
            </span>
          </span>
          <span style="white-space:nowrap;">
            ${delivered_
              ? `<button data-abm-requeue="${i}" style="border:1px solid #ccc;background:#fff;border-radius:6px;padding:2px 8px;cursor:pointer;font-size:11px;">Mark undelivered</button>`
              : `<button data-abm-retry="${i}" style="border:1px solid #1e8e3e;background:#e6f4ea;color:#1e8e3e;border-radius:6px;padding:2px 8px;cursor:pointer;font-size:11px;margin-right:4px;">Retry</button>
                 <button data-abm-markok="${i}" style="border:1px solid #ccc;background:#fff;border-radius:6px;padding:2px 8px;cursor:pointer;font-size:11px;">Mark delivered</button>`}
          </span>
        </div>
        <div style="color:#666;margin-top:3px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">
          ${_abmEsc_((r.messageText || '').slice(0, 90))}</div>
      </div>`;
    body.innerHTML = `
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
        <span style="font-weight:700;color:#c0392b;">Undelivered (${undelivered.length})</span>
        ${undelivered.length ? `<label style="display:flex;align-items:center;gap:4px;font-size:11px;color:#888;cursor:pointer;user-select:none;">
          <input type="checkbox" id="sp-abm-sel-all" style="cursor:pointer;margin:0;"/>select all</label>` : ''}
      </div>
      ${undelivered.length ? undelivered.map((r, i) => rowHtml(r, i, false)).join('') : '<div style="color:#888;padding:4px 0 8px;">None 🎉</div>'}
      <div style="font-weight:700;color:#1e8e3e;margin:12px 0 4px;">Delivered (recent ${delivered.length})</div>
      ${delivered.length ? delivered.map((r, i) => rowHtml(r, i, true)).join('') : '<div style="color:#888;padding:4px 0;">None yet</div>'}`;
    // Header "Retry selected" reads the current selection at click time —
    // stash the undelivered array the checkbox indices refer to.
    panel._abmUndelivered = undelivered;
    const retrySelBtn = panel.querySelector('#sp-abm-retry-sel');
    const updateSelBtn_ = () => {
      if (!retrySelBtn) return;
      const n = body.querySelectorAll('[data-abm-sel]:checked').length;
      retrySelBtn.disabled = !n;
      retrySelBtn.textContent = n ? `Retry selected (${n})` : 'Retry selected';
    };
    body.querySelectorAll('[data-abm-sel]').forEach(cb => cb.onchange = updateSelBtn_);
    const selAll = body.querySelector('#sp-abm-sel-all');
    if (selAll) selAll.onchange = () => {
      body.querySelectorAll('[data-abm-sel]').forEach(cb => { cb.checked = selAll.checked; });
      updateSelBtn_();
    };
    updateSelBtn_();
    body.querySelectorAll('[data-abm-retry]').forEach(btn => btn.onclick = async () => {
      const item = undelivered[Number(btn.dataset.abmRetry)];
      btn.disabled = true; btn.textContent = '…';
      const ok = await trySendPendingItem_(item);
      abmToast_(ok ? `✓ Ticket #${item.ticketId} reply delivered` : `✗ Ticket #${item.ticketId} retry failed — check the Seller Central login for ${item.marketplace}`, ok, true);
      await renderAbmPanelBody_(panel);
      updateAbmBadge_();
    });
    body.querySelectorAll('[data-abm-markok]').forEach(btn => btn.onclick = async () => {
      const item = undelivered[Number(btn.dataset.abmMarkok)];
      btn.disabled = true;
      await postAbmStatus_(item.relayKey, item.ticketId, 'success');
      await renderAbmPanelBody_(panel);
      updateAbmBadge_();
    });
    body.querySelectorAll('[data-abm-requeue]').forEach(btn => btn.onclick = async () => {
      const item = delivered[Number(btn.dataset.abmRequeue)];
      btn.disabled = true;
      await postAbmStatus_(item.relayKey, item.ticketId, 'failed');
      _abmDismissedSig_ = null; // re-queuing is an explicit new problem — un-dismiss the badge
      await renderAbmPanelBody_(panel);
      updateAbmBadge_();
    });
  }

  // Atomic claim before sending — the backend flips the row to 'sending' only
  // for whichever browser wins, so two agents' sweeps can't double-send the
  // same reply to a real buyer. Returns true only if THIS browser got it.
  async function claimAbmRelay_(relayKey) {
    if (!relayKey) return true; // legacy rows without a key: fall back to old behavior
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST', url: GAS_URL, headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify({ action: 'claimAbmRelay', relayKey }),
        timeout: 15000,
        onload: res => { try { resolve(!!JSON.parse(res.responseText).claimed); } catch (_) { resolve(false); } },
        onerror: () => resolve(false), ontimeout: () => resolve(false),
      });
    });
  }

  // Cheap deterministic hash — not cryptographic, just needs to make two
  // near-identical strings collide and two different ones not. Used to key
  // the initial-send claim below by ticket+content, not just ticket, so a
  // genuinely different second reply on the same ticket is never blocked.
  function _abmSimpleHash_(s) {
    let h = 0;
    for (let i = 0; i < s.length; i++) h = (Math.imul(31, h) + s.charCodeAt(i)) | 0;
    return (h >>> 0).toString(36);
  }

  // Mirrors the GAS-side normalizeForDedup_ (ABM_TicketMerge/Code.js) exactly,
  // so the two independent dedup layers agree on what counts as "the same
  // message". Strips HTML entities, keeps only letters/digits, lowercases —
  // aggressive on purpose, so cosmetic differences never defeat the dedup.
  function _normalizeAbmTextForDedup_(s) {
    return String(s || '')
      .replace(/&[a-z]+;|&#\d+;/gi, ' ')
      .replace(/[^\p{L}\p{N}]+/gu, '')
      .toLowerCase();
  }

  // Guards the INITIAL send (relayAbmReply_), a gap claimAbmRelay_ above
  // never covered — that one only stops two browsers double-sending an
  // ALREADY-QUEUED retry row, but nothing stopped the live send itself from
  // firing twice for one real reply in the first place. Confirmed live via
  // ABM_Relay_Log: 28+ tickets since 2026-07-28 show the identical message
  // successfully relayed twice within seconds to minutes of each other, both
  // from this path (not the retry sweep) — e.g. ticket #1000157277, 3
  // seconds apart, only one real Zendesk comment ever existed. Likely
  // triggered by a duplicate/retried Zendesk network call or a second open
  // tab/session on the same ticket catching a live-sync echo — exact cause
  // wasn't pinned down, but any true duplicate firing carries identical
  // ticketId+htmlBody, so keying the claim on both (via the hash above)
  // closes it regardless of the exact mechanism.
  // Fails OPEN (treats a network error/timeout as claimed) — this guard's
  // job is to stop a rare double-fire, not to gate whether a real reply
  // ever reaches the customer; a skipped send from a flaky claim-check
  // would be strictly worse than the duplicate it's meant to prevent. Only
  // an explicit `claimed:false` from the backend (another invocation
  // already holds it) actually blocks the send.
  function claimAbmSend_(key) {
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST', url: GAS_URL, headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify({ action: 'claimAbmSend', key }),
        timeout: 8000,
        onload: res => { try { resolve(JSON.parse(res.responseText).claimed !== false); } catch (_) { resolve(true); } },
        onerror: () => resolve(true), ontimeout: () => resolve(true),
      });
    });
  }

  // Single-item resend, shared by the background sweep and the panel's per-row
  // Retry button. Claims first (no double-send), relays the reply's own comment
  // attachments when it came from reconciliation (item.commentId), and resets
  // the row to 'failed' on a failed send so it's retryable and stays visible
  // instead of stuck on 'sending'.
  async function trySendPendingItem_(item) {
    if (!item || !item.caseId || !item.marketplace || !item.messageText) return false;
    if (!(await claimAbmRelay_(item.relayKey))) return false; // another browser owns it
    // Stage-specific `why` — 'resend attempt failed' used to cover all four
    // failure points identically, which made the log useless for telling a
    // logged-out marketplace (hundreds of no-session sweep ticks from
    // browsers without a UK/IN login) apart from a genuine Amazon rejection.
    const fail = why => {
      logAbmRelayResult_(item.relayKey, item.ticketId, item.caseId, item.marketplace, 'failed', (Number(item.attempts) || 1) + 1, `resend failed: ${why}`, item.messageText, item.commentId);
      return false;
    };
    // item.marketplace is the RAW marketplace as originally logged (e.g.
    // "amazon.ie") — relayAbmReply_ itself already redirects EU non-DE
    // domains to DE before logging (see EU_SC_REDIRECT above), so a row from
    // an up-to-date browser normally already says "amazon.de" here. But this
    // retry path had no redirect of its own, so a row logged before the
    // EU_SC_REDIRECT set covered a given marketplace (or by a stale
    // not-yet-updated browser) got permanently stuck retrying against a
    // domain no agent is ever logged into — confirmed live on ticket
    // #1000157235 (183 failed attempts, all "no SC login session for
    // amazon.ie", when every agent is logged into DE not IE). Applying the
    // same redirect here — not just at initial-send time — makes retries
    // self-heal instead of repeating the same wrong-domain failure forever.
    const scDomain = EU_SC_REDIRECT.has(item.marketplace) ? 'amazon.de' : item.marketplace;
    const base = scBaseUrl_(scDomain);
    const fresh = await resolveScCaseDirect_(base, item.caseId).catch(() => null);
    if (!fresh) return fail(`case lookup failed — likely no SC login session for ${scDomain} in this browser`);
    const resourceVersion = await fetchScResourceVersion_(base).catch(() => null);
    if (!resourceVersion) return fail('resource-version fetch failed');
    let attachmentRefs = [];
    if (item.commentId) {
      const zdAtts = await fetchCommentAttachmentsById_(item.ticketId, item.commentId).catch(() => []);
      attachmentRefs = await relayAbmAttachmentList_(base, item.caseId, fresh.marketplaceId, fresh.csrfToken, resourceVersion, zdAtts).catch(() => []);
    }
    const sent = await sendScMessage_(base, item.caseId, fresh.marketplaceId, null, '', fresh.csrfToken, resourceVersion, item.messageText, attachmentRefs);
    if (!sent.ok) return fail(`send rejected (${sent.why})`);
    logAbmRelayResult_(item.relayKey, item.ticketId, item.caseId, item.marketplace, 'success', (Number(item.attempts) || 1) + 1, '', item.messageText, item.commentId);
    return true;
  }

  // ── Seller Central login-session check ──────────────────────────────────
  // The ABM relay (and the retry sweep below) can only deliver to a
  // marketplace whose SC domain has a live cookie session in THIS browser —
  // sessions the GAS backend can never see, so this must run client-side.
  // Probing every send-domain individually (not one per "region"): the relay
  // routes EU-proper tickets through amazon.de but UK stays on its own
  // amazon.co.uk cookie domain, and the relay log showed browsers holding a
  // DE session while UK sends failed for hundreds of sweep ticks. If SC's
  // SSO does silently cover a sibling domain, its probe just passes and no
  // alert is ever shown — probing separately costs nothing when true.
  const SC_SESSION_DOMAINS = [
    { label: 'EU (DE)', domain: 'amazon.de'    },
    { label: 'UK',      domain: 'amazon.co.uk' },
    { label: 'IN',      domain: 'amazon.in'    },
    { label: 'JP',      domain: 'amazon.co.jp' },
  ];
  const SC_PROBE_OK_LS_KEY = 'gcx_sc_probe_ok_ts';
  const SC_PROBE_OK_TTL_MS = 10 * 60 * 1000;
  let _scLoginBannerDismissed = false;

  // Logged-in ⇢ the messaging API answers JSON; logged-out ⇢ redirect to
  // /ap/signin (HTML). Network error/timeout counts as "missing" too — the
  // banner is dismissible and Recheck is one click, so a rare false alarm
  // beats silently hiding a real logged-out marketplace.
  function probeScSession_(domain) {
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: `${scBaseUrl_(domain)}/messaging/api/global/cases?filterList=ACTION_NEEDED&sortDirection=ASCENDING`,
        headers: { Accept: 'application/json' },
        timeout: 10000,
        onload: res => {
          let ok = !!res && res.status === 200 && !/\/ap\/signin/.test(res.finalUrl || '');
          if (ok) { try { JSON.parse(res.responseText); } catch (_) { ok = false; } }
          resolve(ok);
        },
        onerror:   () => resolve(false),
        ontimeout: () => resolve(false),
      });
    });
  }

  async function checkScSessions_(force) {
    if (!force) {
      // Recently confirmed all-ok (any tab) — skip the probes. A "missing"
      // result is never cached, so a logged-out state keeps being re-checked.
      const okTs = Number(localStorage.getItem(SC_PROBE_OK_LS_KEY) || 0);
      if (okTs && Date.now() - okTs < SC_PROBE_OK_TTL_MS) return;
    }
    const results = await Promise.all(SC_SESSION_DOMAINS.map(d => probeScSession_(d.domain)));
    const missing = SC_SESSION_DOMAINS.filter((_, i) => !results[i]);
    if (!missing.length) {
      try { localStorage.setItem(SC_PROBE_OK_LS_KEY, String(Date.now())); } catch (_) {}
      document.getElementById('sp-sc-login-banner')?.remove();
      _scLoginBannerDismissed = false; // next real gap should alert again
      return;
    }
    logStep_(`SC login check: missing session for ${missing.map(m => m.label).join(', ')}`);
    showScLoginBanner_(missing);
  }

  function showScLoginBanner_(missing) {
    if (_scLoginBannerDismissed) return;
    document.getElementById('sp-sc-login-banner')?.remove();
    const banner = document.createElement('div');
    banner.id = 'sp-sc-login-banner';
    banner.style.cssText = `position:fixed;top:16px;left:50%;transform:translateX(-50%);
      z-index:2147483647;background:#b3541e;color:#fff;padding:10px 14px;border-radius:10px;
      font:600 13px -apple-system,sans-serif;box-shadow:0 4px 16px rgba(0,0,0,.3);
      display:flex;align-items:center;gap:10px;max-width:calc(100vw - 32px);flex-wrap:wrap;`;
    const links = missing.map(m =>
      `<a href="${scBaseUrl_(m.domain)}/messaging/inbox" target="_blank"
          style="color:#ffe6a7;font-weight:700;text-decoration:underline;margin-right:6px;">${m.label} 로그인</a>`).join('');
    banner.innerHTML = `
      <span>⚠ ABM 자동 전송을 위해 Seller Central 로그인이 필요합니다:</span>
      <span>${links}</span>
      <button id="sp-sc-login-recheck" style="border:1px solid rgba(255,255,255,.6);background:transparent;
        color:#fff;border-radius:6px;padding:3px 10px;cursor:pointer;font:600 12px -apple-system,sans-serif;">Recheck</button>
      <button id="sp-sc-login-dismiss" style="border:none;background:transparent;color:#fff;
        cursor:pointer;font-size:15px;line-height:1;padding:2px 4px;">✕</button>`;
    document.body.appendChild(banner);
    banner.querySelector('#sp-sc-login-recheck').onclick = async ev => {
      ev.target.disabled = true; ev.target.textContent = '…';
      await checkScSessions_(true);
      // checkScSessions_ re-renders (still missing) or removes (all ok) the banner.
      const btn = document.querySelector('#sp-sc-login-recheck');
      if (btn) { btn.disabled = false; btn.textContent = 'Recheck'; }
    };
    banner.querySelector('#sp-sc-login-dismiss').onclick = () => {
      _scLoginBannerDismissed = true; // until next full page load
      banner.remove();
    };
  }

  // Runs in EVERY agent's browser on every Zendesk page load (plus a 5-minute
  // interval while the tab stays open), not just the one that originated a
  // failure — since these SC endpoints need a live, cookie-authenticated
  // Seller Central session, there's no single always-on backend that could
  // retry these independently. Instead, whichever agent's browser happens to
  // have a valid session for the relevant marketplace opportunistically picks
  // up and resends any still-pending entry from the shared log. A case this
  // browser can't currently reach (wrong/no SC session for that marketplace)
  // is silently left for another agent's browser to pick up later.
  async function sweepAbmRelayFailures_() {
    const res = await gmGet_(`${GAS_URL}?action=abmRelayPending&clientVersion=${CURRENT_VERSION}`);
    if (!res || res.status !== 200) return;
    let pending;
    try { pending = JSON.parse(res.responseText).pending || []; } catch (_) { return; }
    let recovered = 0;
    for (const item of pending) {
      if (await trySendPendingItem_(item)) recovered++;
    }
    updateAbmBadge_();
    if (recovered) abmToast_(`✓ Recovered ${recovered} previously-failed ABM ${recovered === 1 ? 'reply' : 'replies'}`, true);
  }

  // ── Auto-fill: PUT all fields to Zendesk API, fill text fields in DOM ────

  function autoFillTicket(panel) {
    const ticketId = location.pathname.match(/\/tickets\/(\d+)/)?.[1];
    if (!ticketId || !lastOrderData) return;

    const btn = panel.querySelector('#sp-autofill-btn');
    if (btn) { btn.disabled = true; btn.textContent = 'Preparing…'; }
    setFillStatus(panel, '');

    const o  = lastOrderData.order   || {};
    const ad = lastOrderData.address || {};
    const b  = lastOrderData.buyer   || {};
    const p  = lastProductData || {};

    const orderId   = panel.querySelector('#sp-order-input')?.value.trim() || '';
    const panelAsin = panel.querySelector('#sp-asin-input')?.value.trim()  || '';
    const itemAsins  = (lastOrderData.items || []).map(i => i.ASIN).filter(Boolean);
    const asinValue  = itemAsins.length ? itemAsins.join(', ') : panelAsin;
    const rawSellerSku = lastOrderData.items?.[0]?.SellerSKU || '';
    // Sheet SKU (p.SKU) always takes priority. Only accept rawSellerSku when it matches the
    // Spigen SKU pattern (3 uppercase letters + 5 digits, e.g. "ACH06437", "ACS06557PAN").
    // Rejects barcodes ("8809613760408"), model names ("PE2213IN 35w"), and non-Spigen codes ("PE2213IN").
    const itemSku = p.SKU || (!/^[A-Z]{3}\d{5}/.test(rawSellerSku) ? '' : rawSellerSku) || '';
    // '✅전체 주문'/'❎전체 환불' count total ITEMS purchased/refunded, not order
    // count — a customer with 3 orders totaling 8 items shows 8. Falls back to
    // order-based counts if SC item-quantity data isn't available yet.
    const totalPurchases = lastOrderData.totalItemsPurchased ?? lastOrderData.totalPurchases ?? lastOrderData.orderCount;
    const totalRefunds   = lastOrderData.totalItemsRefunded  ?? lastOrderData.totalRefunds;
    // The two tagger fields use DIFFERENT option tags (verified via
    // /api/v2/ticket_fields): 전체 주문 = q1..q50 (no q0), 전체 환불 = r0..r20.
    // Refunds used to be sent as `q2` etc. — not a valid option for that
    // field, so Zendesk silently dropped it and 전체 환불 never got applied.
    const purchasesVal   = totalPurchases != null && totalPurchases > 0 ? `q${Math.min(totalPurchases, 50)}` : null;
    const refundsVal     = totalRefunds   != null && totalRefunds  >= 0 ? `r${Math.min(totalRefunds,   20)}` : null;
    const buyerName    = b.BuyerName || o.BuyerInfo?.BuyerName || ad.Name || getTicketRequesterNameFromDom_() || '';
    const orderTotal   = o.OrderTotal ? `${o.OrderTotal.Amount} ${o.OrderTotal.CurrencyCode}` : '';
    const purchaseDateIso = purchaseDateLocal_(o.PurchaseDate, ad.CountryCode);
    const amz = lastAmazonProduct || {};
    const daebunryu  = p['대분류']     || amz['대분류']     || '';
    const saengsan   = p['생산업체']   || amz['생산업체']   || '';
    const originInfo = p['원산지정보'] || amz['원산지정보'] || '';
    const deviceLabel    = p['기종명'] || '';
    const productLabel   = p['모델명']  || '';
    const ticketText     = getCachedBodyText_();
    // PAN EU detection must check the raw SP-API SellerSKU (e.g. "ACS06557PAN"), NOT the
    // sheet SKU — the sheet stores the base SKU without the PAN/EUP suffix.
    const skuHasPan      = /pan|eup/i.test(rawSellerSku);
    const fulfillChannel = o.FulfillmentChannel || '';
    const brandTag       = brandFromDaebunryu(p['대분류'] || '');
    const pop            = salesChannelToPOP(o.SalesChannel);

    // DOM text-field rows — read current values for "before" column
    const DOM_DEFS = [
      { label: 'Order ID',           zdId: ZD.ORDER_ID,     after: orderId,           apiVal: orderId },
      { label: 'ASIN',               zdId: ZD.ASIN,         after: asinValue,          apiVal: asinValue },
      { label: '문의SKU',            zdId: ZD.SKU,          after: itemSku,            apiVal: itemSku },
      { label: 'Customer Full Name', zdId: ZD.CUST_NAME,    after: buyerName,          apiVal: buyerName },
      { label: 'Purchase Date',      zdId: ZD.PURCHASE_DATE,after: purchaseDateIso,    apiVal: purchaseDateIso, isDate: true },
      { label: 'Order Status',       zdId: ZD.ORDER_STATUS, after: o.OrderStatus||'',  apiVal: o.OrderStatus||'' },
      { label: 'Order Total',        zdId: ZD.ORDER_TOTAL,  after: orderTotal,         apiVal: orderTotal },
      { label: 'Delivery Level',     zdId: ZD.DELIVERY_LVL, after: o.ShipmentServiceLevelCategory||'', apiVal: o.ShipmentServiceLevelCategory||'' },
      { label: '대분류',             zdId: ZD.DAEBUNRYU,    after: daebunryu,          apiVal: daebunryu },
      { label: '생산업체',           zdId: ZD.SAENGSAN,     after: saengsan,           apiVal: saengsan },
      { label: '원산지정보',         zdId: ZD.ORIGIN_INFO,  after: originInfo,         apiVal: originInfo },
    ];
    const textRows = DOM_DEFS.map(d => ({
      label: resolveZdLabel_(d.label) || d.label, before: readZdInput_(d.label), after: d.after,
      dom: d, // carry definition for execution phase
    }));

    // Async-resolved rows (Device, Product Name, Fulfillment, Photo, API-only fields)
    let resolvedDevice   = null; // { val, name, opts }
    let resolvedProduct  = null; // { val, name }
    let resolvedFulfill  = null; // { val, name }
    let fulfillOpts      = [];   // all fulfillment opts — for before display name
    let resolvedHasPhoto = null;
    let currentCfMap     = {};   // fieldId → raw value from current ticket
    let ticketSubject    = '';   // captured from ticket JSON — used for invoice detection
    // Customer Full Name: starts as whatever SP-API gave us (may be empty —
    // BuyerName PII is often unavailable); patched below via DOM scrape if so.
    let resolvedRequesterName = buyerName;

    const BRAND_TAG_LABEL = {
      spigen_case_: 'CASE', spigen_sp_: 'SP', spigen_sda_: 'SDA',
      'spigen_pacc._': 'PAcc.', spigen_new_biz_: 'Newbiz', 'n/a': 'n/a',
    };
    // Full option lists for the Confirm Auto-Fill popup's "After" dropdowns —
    // every tagger/dropdown-backed ZD field gets its full, real choice list so
    // the agent can correct a wrong auto-pick, not just accept/reject it.
    const COUNTRY_OPTS  = Object.entries(COUNTRY_MAP).map(([iso, value]) => ({ value, name: iso }));
    const POP_OPTS = [
      { value: 'amazon_eu',              name: 'Amazon EU' },
      { value: 'amazon_united_kingdom',  name: 'Amazon UK' },
      { value: 'amazon_japan',           name: 'Amazon Japan' },
      { value: 'amazon_singapore',       name: 'Amazon Singapore' },
      { value: 'amazon_india',           name: 'Amazon India' },
      { value: 'others',                 name: 'Others' },
    ];
    const BRAND_OPTS = Object.entries(BRAND_TAG_LABEL).map(([value, name]) => ({ value, name }));
    const PHOTO_OPTS = [{ value: 'yes', name: 'Yes' }, { value: 'no', name: 'No' }];
    const Q_OPTS = Array.from({ length: 50 }, (_, i) => ({ value: `q${i + 1}`, name: `q${i + 1}` }));
    const R_OPTS = Array.from({ length: 21 }, (_, i) => ({ value: `r${i}`,     name: `r${i}` }));

    function buildAndShow() {
      // Invoice ticket override: if subject contains "invoice", force Device to the invoice option
      if (/invoice/i.test(ticketSubject) && resolvedDevice?.opts) {
        const invoiceOpt = resolvedDevice.opts.find(o => /invoice/i.test(o.name));
        if (invoiceOpt) resolvedDevice = { val: invoiceOpt.value, name: invoiceOpt.name, opts: resolvedDevice.opts };
      }

      // Refresh textRows "before" from ZD API (currentCfMap) — more reliable than DOM snapshot
      for (const row of textRows) {
        const cf = currentCfMap[row.dom.zdId];
        if (cf !== undefined) row.before = cf;
      }

      // Patch in the async-resolved requester name (only differs from the
      // DOM_DEFS snapshot when SP-API had no BuyerName and the DOM fallback
      // had to open the Customer context panel to find it).
      const custRow = textRows.find(r => r.dom?.zdId === ZD.CUST_NAME);
      if (custRow && resolvedRequesterName) {
        custRow.after = resolvedRequesterName;
        custRow.dom.after = resolvedRequesterName;
        custRow.dom.apiVal = resolvedRequesterName;
      }

      // API-only rows (no DOM text field counterpart)
      const apiOnlyRows = [
        { label: 'Device*', zdId: ZD.DEVICE, before: resolvedDevice?.opts?.find(o => o.value === currentCfMap[ZD.DEVICE])?.name || currentCfMap[ZD.DEVICE] || '',
          after: resolvedDevice?.name || '', api: resolvedDevice?.val ? { id: ZD.DEVICE, value: resolvedDevice.val } : null,
          candidates: resolvedDevice?.candidates || null, opts: resolvedDevice?.opts || [] },
        { label: 'Product Name *', zdId: ZD.PRODUCT_NAME, before: '', // resolved below using product opts
          after: resolvedProduct?.name || '', api: resolvedProduct?.val ? { id: ZD.PRODUCT_NAME, value: resolvedProduct.val } : null,
          candidates: resolvedProduct?.candidates || null, opts: resolvedProduct?.opts || [] },
        { label: 'Country*', zdId: ZD.COUNTRY,      before: currentCfMap[ZD.COUNTRY]     || '', after: COUNTRY_MAP[ad.CountryCode] || '', api: COUNTRY_MAP[ad.CountryCode] ? { id: ZD.COUNTRY, value: COUNTRY_MAP[ad.CountryCode] } : null, opts: COUNTRY_OPTS },
        { label: 'Point of Purchase', zdId: ZD.POINT_OF_PUR, before: currentCfMap[ZD.POINT_OF_PUR] || '', after: pop || '', api: pop ? { id: ZD.POINT_OF_PUR, value: pop } : null, opts: POP_OPTS },
        { label: 'Amazon Fulfillment Methods*', zdId: ZD.FULFILLMENT, before: fulfillOpts.find(o => o.value === currentCfMap[ZD.FULFILLMENT])?.name || currentCfMap[ZD.FULFILLMENT] || '', after: resolvedFulfill?.name || '', api: resolvedFulfill?.val ? { id: ZD.FULFILLMENT, value: resolvedFulfill.val } : null, opts: fulfillOpts },
        { label: 'Brand(상세)*', zdId: ZD.BRAND_DETAIL, before: BRAND_TAG_LABEL[currentCfMap[ZD.BRAND_DETAIL]] || currentCfMap[ZD.BRAND_DETAIL] || '', after: brandTag || '', api: brandTag ? { id: ZD.BRAND_DETAIL, value: brandTag } : null, opts: BRAND_OPTS },
        { label: '✅전체 주문 (Product Issue, 아크테크X)*', zdId: ZD.TOTAL_ORDERS, before: currentCfMap[ZD.TOTAL_ORDERS]  || '', after: purchasesVal || '', api: purchasesVal ? { id: ZD.TOTAL_ORDERS, value: purchasesVal } : null, opts: Q_OPTS },
        { label: '❎전체 환불*', zdId: ZD.TOTAL_REFUNDS, before: currentCfMap[ZD.TOTAL_REFUNDS] || '', after: refundsVal || '', api: refundsVal ? { id: ZD.TOTAL_REFUNDS, value: refundsVal } : null, opts: R_OPTS },
        { label: '❗사진/영상 유무❗*', zdId: ZD.PHOTO_EXIST, before: currentCfMap[ZD.PHOTO_EXIST] || '', after: resolvedHasPhoto != null ? (resolvedHasPhoto ? 'yes' : 'no') : '', api: resolvedHasPhoto != null ? { id: ZD.PHOTO_EXIST, value: resolvedHasPhoto ? 'yes' : 'no' } : null, opts: PHOTO_OPTS },
      ];

      // Resolve Product Name "before" using product opts
      const pnRow = apiOnlyRows.find(r => r.label === 'Product Name *');
      if (pnRow && resolvedProduct?.opts) {
        pnRow.before = resolvedProduct.opts.find(o => o.value === currentCfMap[ZD.PRODUCT_NAME])?.name?.replace(/^[^_]+_/, '') || currentCfMap[ZD.PRODUCT_NAME] || '';
      }

      const allRows = [...textRows, ...apiOnlyRows];

      if (btn) { btn.disabled = false; btn.textContent = 'Auto-Fill Form'; }

      showFillConfirm_(allRows, selectedRows => {
        _gcrFilledThisTicket = true; // track that GCX Reply filled this ticket's fields
        if (btn) { btn.disabled = true; btn.textContent = 'Filling…'; }
        const _t0 = performance.now();
        // Only clear fields we're about to refill (wipes stale React state at fill-time
        // for those specific fields). Clearing every tracked field unconditionally used
        // to wipe out pre-existing agent-entered values — e.g. Order ID/ASIN/SKU — any
        // time GCX Reply itself had no replacement value for that ticket, even though
        // those rows are correctly excluded from selectedRows (disabled in the confirm
        // modal when there's nothing new to fill).
        clearAllZdFields_(selectedRows.filter(r => r.dom).map(r => r.dom.label));
        const _t1 = performance.now();

        // DOM fills first
        let dispatchEsc = false;
        for (const r of selectedRows) {
          if (!r.dom) continue;
          fillZdInput(r.dom.label, r.dom.after);
          if (r.dom.isDate && r.dom.after) dispatchEsc = true;
        }
        if (dispatchEsc) {
          setTimeout(() => document.dispatchEvent(
            new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })
          ), 150);
        }
        const _t2 = performance.now();

        // Build API payload from selected rows
        const af = [];
        for (const r of selectedRows) {
          if (r.api?.value) { af.push(r.api); continue; }
          if (!r.dom) continue;
          // DOM rows also write via API
          const v = r.dom.label === 'Purchase Date' ? r.dom.apiVal : r.dom.after;
          if (v) af.push({ id: r.dom.zdId, value: v });
        }
        logStep_(`Fill timing — clear: ${(_t1-_t0).toFixed(0)}ms, DOM fill: ${(_t2-_t1).toFixed(0)}ms, fields: ${selectedRows.length}, PUT starting…`);

        putZdTicket(ticketId, af, btn, panel, _t2);
      }, () => {
        if (btn) { btn.disabled = false; btn.textContent = 'Auto-Fill Form'; }
      });
    }

    // Async fetches — device opts always fetched (invoice detection needs opts even without product).
    // Each source resolves its own Promise once its side-effect (resolvedDevice
    // etc.) is set; Promise.all below calls buildAndShow() exactly once, all
    // sources accounted for by construction — no manual counter to keep in
    // sync by hand when a source is added or removed (a real bug risk: a
    // previous change here required remembering to bump a shared count).
    const deviceReady = new Promise(resolve => {
      fetchZdFieldOptsCached(ZD.DEVICE, opts => {
        if (deviceLabel) {
          const cands = topDeviceCandidates_(opts, deviceLabel, ticketText);
          if (cands.length === 0) {
            resolvedDevice = { val: null, name: '', opts };
          } else if (cands.length === 1) {
            resolvedDevice = { val: cands[0].val, name: cands[0].name, opts };
          } else {
            resolvedDevice = { val: cands[0].val, name: cands[0].name, opts, candidates: cands };
          }
        } else {
          resolvedDevice = { val: null, name: '', opts };
        }
        resolve();
      });
    });
    // Always fetch (like Device above) — even with no productLabel to match
    // against, the field's full option list still needs to reach the Confirm
    // Auto-Fill popup so it renders as a dropdown instead of falling back to
    // a plain text input.
    const productReady = new Promise(resolve => {
      fetchZdFieldOptsCached(ZD.PRODUCT_NAME, opts => {
        const cands = productLabel ? topCandidateOpts_(opts, productLabel, true) : [];
        if (cands.length === 0) {
          resolvedProduct = { val: null, name: '', opts };
        } else if (cands.length === 1) {
          resolvedProduct = { val: cands[0].val, name: cands[0].displayName, opts };
        } else {
          resolvedProduct = { val: cands[0].val, name: cands[0].displayName, opts, candidates: cands };
        }
        resolve();
      });
    });
    const fulfillmentReady = new Promise(resolve => {
      fetchZdFieldOptsCached(ZD.FULFILLMENT, opts => {
        fulfillOpts = opts;
        let fv = null;
        if (skuHasPan) { const o2 = opts.find(x => /pan\s*eu/i.test(x.name)); if (o2) fv = o2.value; }
        if (!fv && fulfillChannel) {
          const kw = fulfillChannel === 'AFN' ? 'fba' : fulfillChannel === 'MFN' ? 'merchant' : null;
          if (kw) { const o2 = opts.find(x => x.name.toLowerCase().startsWith(kw)); if (o2) fv = o2.value; }
        }
        resolvedFulfill = fv ? { val: fv, name: opts.find(o => o.value === fv)?.name || fv } : null;
        resolve();
      });
    });
    const commentsReady = new Promise(resolve => {
      fetchTicketComments(ticketId, hasPhoto => { resolvedHasPhoto = hasPhoto; resolve(); });
    });
    const ticketJsonReady = new Promise(resolve => {
      fetchTicketJson_(ticketId, t => {
        if (t) {
          for (const f of (t.custom_fields || [])) currentCfMap[f.id] = f.value || '';
          ticketSubject = t.subject || '';
        }
        resolve();
      });
    });
    // Customer Full Name DOM fallback — no-ops instantly (no panel-open
    // latency) when SP-API already provided BuyerName.
    const requesterNameReady = buyerName ? Promise.resolve() : ensureCustomerContextPanelOpen_().then(async switchedTab => {
      // The tab reporting "open" (aria-pressed=true) doesn't mean its
      // content has actually rendered yet — confirmed live for the Notes
      // textarea (same lazy-rendered card), so poll here too instead of
      // reading once immediately after the tab-open promise resolves.
      resolvedRequesterName = await pollUntil_(getTicketRequesterNameFromDom_);
      // Restore the Apps tab (where GCX Reply itself lives) if we're the
      // one who switched away from it to read the name.
      if (switchedTab) {
        const appsBtn = document.querySelector('[data-test-id="omnipanel-selector-item-apps"]');
        if (appsBtn && appsBtn.getAttribute('aria-pressed') !== 'true') appsBtn.click();
      }
    });

    Promise.all([deviceReady, productReady, fulfillmentReady, commentsReady, ticketJsonReady, requesterNameReady])
      .then(buildAndShow);
  }

  function putZdTicket(ticketId, af, btn, panel, _tStart) {
    if (!af.length) {
      if (btn) { btn.disabled = false; btn.textContent = 'Auto-Fill Form'; }
      setFillStatus(panel, 'Nothing to fill.');
      return;
    }
    const csrf = document.querySelector('meta[name="csrf-token"]')?.content || '';
    const _tPutStart = performance.now();
    GM_xmlhttpRequest({
      method:  'PUT',
      url:     `https://spigenhelp.zendesk.com/api/v2/tickets/${ticketId}.json`,
      headers: { 'Content-Type': 'application/json', 'X-CSRF-Token': csrf },
      data:    JSON.stringify({ ticket: { custom_fields: af } }),
      onload(res) {
        if (btn) { btn.disabled = false; btn.textContent = 'Auto-Fill Form'; }
        setFillStatus(panel, res.status === 200 ? `✓ ${af.length} fields saved` : `API error ${res.status}`);
        const putMs = performance.now() - _tPutStart;
        const totalMs = _tStart != null ? performance.now() - _tStart : null;
        logStep_(`Fill timing — PUT: ${putMs.toFixed(0)}ms${totalMs != null ? `, total: ${totalMs.toFixed(0)}ms` : ''} (status ${res.status})`);
      },
      onerror() {
        if (btn) { btn.disabled = false; btn.textContent = 'Auto-Fill Form'; }
        setFillStatus(panel, 'Network error');
        logStep_(`Fill timing — PUT: network error after ${(performance.now() - _tPutStart).toFixed(0)}ms`);
      },
    });
  }

  // ── AI 인입사유 ───────────────────────────────────────────────────────────
  function showAiReasonBtn_(category) {
    const container = document.getElementById('sp-ai-reason-result');
    if (!container) return;
    container.innerHTML = `
      <div id="sp-ai-reason-bar">
        <button id="sp-ai-reason-btn"><span class="sp-ai-star">✦</span> AI 인입사유 분석</button>
      </div>`;
    container.querySelector('#sp-ai-reason-btn').addEventListener('click', () => {
      const review = getAiInputText_();
      fetchAiReason_(review, category);
    });
  }

  function fetchAiReason_(review, category) {
    const container = document.getElementById('sp-ai-reason-result');
    if (!container || !review) return;
    const _session = _panelSession;
    container.innerHTML = `
      <div id="sp-ai-reason-bar">
        <button id="sp-ai-reason-btn" disabled><span class="sp-ai-star sp-ai-spin">✦</span> AI 인입사유 분석 중…</button>
      </div>`;
    logStep_('AI 인입사유 분석 중…');
    logStep_('AI 입력텍스트: ' + review.slice(0, 600).replace(/\n+/g, ' / '));
    GM_xmlhttpRequest({
      method:   'GET',
      url:      `${GAS_URL}?action=inferReason&review=${encodeURIComponent(review.slice(0, 2000))}&category=${encodeURIComponent(category || '')}`,
      redirect: 'follow',
      timeout:  35000,
      onload(res) {
        if (_panelSession !== _session || !container.isConnected) return;
        try {
          const data = JSON.parse(res.responseText);
          if (data.error) logStep_(`AI GAS오류: ${data.error}`);
          lastAiReason = data.reason || null;
          renderAiReason_(lastAiReason);
          logStep_(`AI 인입사유: ${lastAiReason || '(결과 없음)'}`);
        } catch (err) {
          renderAiReason_(null);
          logStep_(`AI 인입사유 오류: JSON파싱실패 — ${res.responseText.slice(0, 120)}`);
        }
      },
      onerror(res) {
        if (_panelSession !== _session) return;
        container.innerHTML = '';
        logStep_(`AI 인입사유 오류: network error (status=${res.status})`);
      },
      ontimeout() {
        if (_panelSession !== _session) return;
        container.innerHTML = '';
        logStep_('AI 인입사유 오류: timeout (35s 초과)');
      },
    });
  }

  function renderAiReason_(reason) {
    const container = document.getElementById('sp-ai-reason-result');
    if (!container) return;
    const color = reason ? '#7c3aed' : '#9ca3af';
    const fill  = reason ? '#7c3aed' : '#9ca3af';
    container.innerHTML = `
      <div class="sp-block" data-sp-section="ai_reason">
        <div class="sp-block-title" style="color:${color};border-top:1px solid #e9ebec;">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="${fill}" xmlns="http://www.w3.org/2000/svg">
            <path d="M12 2l1.09 6.26L19 6l-4.26 4.91L21 12l-6.26 1.09L19 19l-4.91-4.26L12 21l-1.09-6.26L5 18l4.26-4.91L3 12l6.26-1.09L5 6l4.91 4.26L12 2z"/>
          </svg>
          AI 인입사유
          <span class="sp-chevron">▾</span>
        </div>
        <div class="sp-block-body">
          ${row('인입사유', reason || '(분류 불가)')}
        </div>
      </div>`;
    container.querySelectorAll('.sp-block-title').forEach(t => {
      t.addEventListener('click', e => { e.stopPropagation(); t.closest('.sp-block').classList.toggle('collapsed'); });
    });
    applySectionState(container);
    // Always expand after a fresh analysis — user just asked for it
    const _aiBlock = container.querySelector('[data-sp-section="ai_reason"]');
    if (_aiBlock) _aiBlock.classList.remove('collapsed');
  }

  // Product sheets (ASIN Master, market sheets) change weekly-to-monthly, so a
  // GAS ?asin= response is cached per ASIN in GM storage for PRODUCT_CACHE_TTL
  // — repeat ASINs render instantly with no GAS round-trip. Only clean JSON
  // without `error` is cached (never HTML error pages / timeouts). The manual
  // Product button and the GAS-not-ready retry bypass the cache (and refresh it).
  const PRODUCT_CACHE_KEY = 'gcx_gas_product_cache';
  const PRODUCT_CACHE_TTL = 6 * 3600 * 1000;
  const PRODUCT_CACHE_MAX = 500;
  function gasProductRequest_(asin, opts, bypassCache) {
    let all = {};
    try { all = GM_getValue(PRODUCT_CACHE_KEY, {}) || {}; } catch (_) {}
    const hit = all[asin];
    if (!bypassCache && hit && Date.now() - hit.ts < PRODUCT_CACHE_TTL) {
      setTimeout(() => opts.onload({ status: 200, responseText: hit.json }), 0);
      return;
    }
    const onload = opts.onload;
    backendRequest_(`asin=${encodeURIComponent(asin)}`, Object.assign({}, opts, {
      onload(res) {
        try {
          const txt = res.responseText || '';
          if (res.status === 200 && !txt.trimStart().startsWith('<') && !JSON.parse(txt).error) {
            const cur = GM_getValue(PRODUCT_CACHE_KEY, {}) || {};
            cur[asin] = { json: txt, ts: Date.now() };
            const keys = Object.keys(cur);
            if (keys.length > PRODUCT_CACHE_MAX) {
              keys.sort((a, b) => cur[a].ts - cur[b].ts).slice(0, keys.length - PRODUCT_CACHE_MAX).forEach(k => delete cur[k]);
            }
            GM_setValue(PRODUCT_CACHE_KEY, cur);
          }
        } catch (_) {}
        onload(res);
      },
    }));
  }

  function renderAllProducts(asins, _retry, _forceExpand, _fresh) {
    const container = document.getElementById('sp-product-result');
    if (!container) return;
    if (!asins.length) { container.innerHTML = ''; return; }
    const _session = _panelSession;
    container.innerHTML = `<div style="font-size:11px;color:#aaa;padding:4px 14px;">Loading product info…</div>`;
    if (!_retry) logStep_(`GAS product lookup: ${asins.join(', ')}`);
    // Start the Amazon /dp/ fetch now, in parallel with GAS — the ASIN Sources
    // block (and the not-in-sheet fallback) always need it afterwards, and
    // fetchAmazonProduct_ hands the same in-flight/cached result to them.
    if (!_retry) asins.forEach(a => fetchAmazonProduct_(a, () => {}));

    let loaded = 0;
    const results = new Array(asins.length).fill(null);

    function done(idx) { if (++loaded === asins.length) finish(); }

    asins.forEach((asin, idx) => {
      gasProductRequest_(asin, {
        method:   'GET',
        url:      `${GAS_URL}?asin=${encodeURIComponent(asin)}`,
        redirect: 'follow',
        timeout:  30000,
        onload(res) {
          if (res.responseText.trimStart().startsWith('<')) {
            results[idx] = { asin, product: null, source: null, marketplaces: [], error: '__html__', allSources: null };
            done(idx);
            return;
          }
          try {
            const data = JSON.parse(res.responseText);
            const mkts = data.marketplaces || [];
            if (data.product && data.productSource !== 'market') {
              // Full data from sheet1 or sheet2 — use directly
              logStep_(`Product: found in ${data.productSource || 'sheet'} (${asin})`);
              results[idx] = { asin, product: data.product, source: data.productSource || 'sheet', marketplaces: mkts, allSources: data.allSources || null };
              done(idx);
            } else if (data.product && data.productSource === 'market') {
              // Partial from country sheet (기종명 col A, 모델명 col B) — merge with Amazon
              logStep_(`Product: market sheet partial (${asin}), fetching Amazon…`);
              const partial = data.product;
              fetchAmazonProduct_(asin, (amazonProduct, amazonUrl) => {
                let merged = partial, src = 'market';
                if (amazonProduct) {
                  merged = Object.assign({}, amazonProduct);
                  if (partial['기종명']) merged['기종명'] = partial['기종명'];
                  if (partial['모델명']) merged['모델명'] = partial['모델명'];
                  src = 'market+amazon';
                }
                logStep_(amazonProduct ? `Product: Amazon merged (${asin})` : `Product: Amazon not found (${asin})`);
                results[idx] = { asin, product: merged, source: src, sourceUrl: amazonUrl, marketplaces: mkts, allSources: data.allSources || null, amazonProduct: amazonProduct || null, amazonUrl };
                done(idx);
              });
            } else {
              // Not in any sheet → fall back to Amazon product page
              logStep_(`Product: not in sheets (${asin}), fetching Amazon…`);
              fetchAmazonProduct_(asin, (amazonProduct, amazonUrl) => {
                logStep_(amazonProduct ? `Product: Amazon (${asin})` : `Product: not found anywhere (${asin})`);
                results[idx] = {
                  asin,
                  product:      amazonProduct,
                  source:       amazonProduct ? 'amazon' : null,
                  sourceUrl:    amazonUrl,
                  marketplaces: mkts,
                  error:        amazonProduct ? null : `${asin} not found in any sheet or Amazon page.`,
                  allSources:   data.allSources || null,
                  amazonProduct: amazonProduct || null,
                  amazonUrl,
                };
                done(idx);
              });
            }
          } catch (err) {
            results[idx] = { asin, product: null, source: null, marketplaces: [], error: 'Parse error: ' + err.message, allSources: null };
            done(idx);
          }
        },
        onerror() {
          results[idx] = { asin, product: null, source: null, error: 'Cannot reach GAS endpoint.', allSources: null };
          done(idx);
        },
      }, _retry || _fresh);
    });

    function sourceBadge_(source, sourceUrl) {
      const MKTSS   = '172fDVw4tu-hgbpV5FShWj4_SAMxeB54-v5BUlVgJUoA';
      const sheet2Url = `https://docs.google.com/spreadsheets/d/${MKTSS}/edit?gid=583143689`;
      const b = (href, label, bg) =>
        `<a href="${esc(href || '#')}" target="_blank" rel="noopener"
          style="font-size:10px;font-weight:normal;background:${bg};color:#fff;
                 padding:1px 6px;border-radius:3px;margin-left:4px;text-decoration:none;">${label}</a>`;
      if (source === 'sheet' || source === 'sheet1') return b(SHEET_URL, 'Sheet', '#34a853');
      if (source === 'sheet2')                        return b(sheet2Url, 'Sheet', '#34a853');
      if (source === 'market')                        return b(sheet2Url, 'Mkt',   '#e67e22');
      if (source === 'market+amazon')                 return b(sheet2Url, 'Mkt', '#e67e22') + b(sourceUrl, 'Amazon', '#FF9900');
      if (source === 'amazon')                        return b(sourceUrl, 'Amazon', '#FF9900');
      return '';
    }

    function finish() {
      if (_panelSession !== _session || !container.isConnected) return;
      if (!_retry && results.every(r => r.error === '__html__')) {
        logStep_('GAS not ready, retrying product lookup…');
        container.innerHTML = `<div style="font-size:11px;color:#aaa;padding:4px 14px;">Retrying…</div>`;
        setTimeout(() => renderAllProducts(asins, true, _forceExpand), 2000);
        return;
      }
      results.forEach(r => { if (r.error === '__html__') r.error = 'GAS error — refresh and try again'; });
      const valid = results.filter(r => r.product);
      // Prefer sheet data for auto-fill (sheet1 > sheet2 > market+amazon > amazon > market)
      lastProductData =
        valid.find(r => r.source === 'sheet' || r.source === 'sheet1' || r.source === 'sheet2')?.product ||
        valid.find(r => r.source === 'market+amazon')?.product ||
        valid[0]?.product || null;
      // Store Amazon product for fallback (대분류/생산업체/원산지정보 may be empty in sheet)
      const amzResult = valid.find(r => r.amazonProduct);
      if (amzResult?.amazonProduct) lastAmazonProduct = amzResult.amazonProduct;
      _productReady = true;  // product lookup finished — enable Auto-Fill Form button
      maybeShowAutoFill(document.getElementById(PANEL_ID));

      const aiCategory = lastProductData?.['대분류'] || '';
      showAiReasonBtn_(aiCategory);

      container.innerHTML = `<div style="padding:0 14px 8px;">${results.map(({ asin, product, source, sourceUrl, error, marketplaces }) => {
        if (!product) {
          const msg = error || `${esc(asin)} not found.`;
          return `<div style="font-size:11px;color:${error ? '#c00' : '#aaa'};padding:4px 0;">${esc(msg)}</div>`;
        }
        const label = asins.length > 1 ? esc(asin) : 'Product Info';
        return `
          <div class="sp-block" style="margin-top:0;" data-sp-section="product_${esc(asin)}">
            <div class="sp-block-title" style="border-top:1px solid #e9ebec;">
              ${label}${sourceBadge_(source, sourceUrl)}
              <span class="sp-chevron">▾</span>
            </div>
            <div class="sp-block-body">
              ${marketplacesRow_(marketplaces)}
              ${SHEET_COLS.map(col => row(col, product[col])).join('')}
            </div>
          </div>`;
      }).join('')}</div>`;

      container.querySelectorAll('.sp-block-title').forEach(t => {
        t.addEventListener('click', e => { e.stopPropagation(); t.closest('.sp-block').classList.toggle('collapsed'); });
      });

      if (_forceExpand) {
        // ASIN-only ticket: apply saved state for non-product sections only;
        // product blocks stay expanded (default) so product info is the focus
        const _c = loadUi().collapsed || {};
        container.querySelectorAll('[data-sp-section]').forEach(block => {
          const key = block.dataset.spSection;
          if (key.startsWith('product_')) return;
          if (!(key in _c)) return;
          if (_c[key]) block.classList.add('collapsed'); else block.classList.remove('collapsed');
        });
      } else {
        applySectionState(container);
      }
      appendSourcesSection_(container, results);
    }
  }

  // ── ASIN source blocks ────────────────────────────────────────────────────

  function buildSourceBlock_(title, linkUrl, product) {
    const link = linkUrl
      ? ` <a href="${esc(linkUrl)}" target="_blank" rel="noopener"
           style="font-size:10px;font-weight:normal;color:#5ba4cf;text-decoration:none;margin-left:4px;">↗</a>`
      : '';
    // Full product page title (Amazon source only) — shown as a "Title" row above SKU.
    const titleRow = product._title
      ? `<div class="sp-row"><span class="sp-label" style="font-size:11.5px;">Title</span><span class="sp-val" style="font-size:11.5px;line-height:1.35;">${esc(product._title)}</span>${COPY_BTN}</div>`
      : '';
    const fields = SHEET_COLS.map(col => {
      const val = product[col];
      if (!val) return '';
      return `<div class="sp-row"><span class="sp-label" style="font-size:11.5px;">${esc(col)}</span><span class="sp-val" style="font-size:11.5px;">${esc(val)}</span>${COPY_BTN}</div>`;
    }).filter(Boolean).join('');
    return `
      <div class="sp-block collapsed" style="margin-top:0;">
        <div class="sp-block-title" style="border-top:1px solid #e9ebec;font-size:11.5px;">
          ${esc(title)}${link}<span class="sp-chevron" style="margin-left:auto;">▾</span>
        </div>
        <div class="sp-block-body">
          ${titleRow}${fields || '<div class="sp-row"><span class="sp-val" style="color:#aaa;font-size:11px;">No data</span></div>'}
        </div>
      </div>`;
  }

  function addCollapseListeners_(el) {
    el.querySelectorAll('.sp-block-title').forEach(t => {
      t.addEventListener('click', e => { e.stopPropagation(); t.closest('.sp-block').classList.toggle('collapsed'); });
    });
  }

  function appendSourcesSection_(container, results) {
    const SHEET1_LINK = 'https://docs.google.com/spreadsheets/d/1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo/edit?gid=0#gid=0';
    const SHEET2_LINK = 'https://docs.google.com/spreadsheets/d/172fDVw4tu-hgbpV5FShWj4_SAMxeB54-v5BUlVgJUoA/edit?gid=716900287';

    const wrap = document.createElement('div');
    wrap.style.cssText = 'padding:0 14px 4px;';

    // Collapsible header — collapsed by default
    const hdr = document.createElement('div');
    hdr.style.cssText = 'font-size:11.5px;font-weight:600;color:#5ba4cf;padding:8px 0 4px;border-top:1px solid #e9ebec;margin-top:4px;cursor:pointer;display:flex;align-items:center;justify-content:space-between;user-select:none;';
    const hdrText = document.createElement('span');
    hdrText.textContent = 'ASIN Sources';
    const hdrChevron = document.createElement('span');
    hdrChevron.textContent = '▸';
    hdrChevron.style.cssText = 'font-size:10px;color:#aaa;';
    hdr.appendChild(hdrText);
    hdr.appendChild(hdrChevron);

    const body = document.createElement('div');
    const _asinSrcCollapsed = (loadUi().collapsed || {})['asin_sources'] !== false;
    body.style.display = _asinSrcCollapsed ? 'none' : '';
    hdrChevron.textContent = _asinSrcCollapsed ? '▸' : '▾';

    hdr.addEventListener('click', () => {
      const open = body.style.display !== 'none';
      body.style.display = open ? 'none' : '';
      hdrChevron.textContent = open ? '▸' : '▾';
      const c = loadUi().collapsed || {};
      c['asin_sources'] = open;
      saveUi({ collapsed: c });
    });

    wrap.appendChild(hdr);
    wrap.appendChild(body);
    container.appendChild(wrap);
    logStep_('Checking ASIN sources…');

    results.forEach(r => {
      if (!r || !r.asin) return;
      const { asin, allSources } = r;

      if (results.length > 1) {
        const lbl = document.createElement('div');
        lbl.style.cssText = 'font-size:11px;font-weight:600;color:#888;padding:2px 0;font-family:monospace;';
        lbl.textContent = asin;
        body.appendChild(lbl);
      }

      // ASIN Master (Sheet1)
      const s1el = document.createElement('div');
      if (allSources === null) {
        s1el.innerHTML = `<div style="font-size:11px;color:#c00;padding:2px 0;">ASIN Master — fetch error</div>`;
        logStep_(`Source: ASIN Master fetch error (${asin})`);
      } else {
        const s1 = allSources.sheet1;
        if (s1) {
          s1el.innerHTML = buildSourceBlock_('✓ ASIN Master', SHEET1_LINK, s1);
          logStep_(`Source: ASIN Master found (${asin})`);
        } else {
          s1el.innerHTML = `<div style="font-size:11px;color:#bbb;padding:2px 0;">✗ ASIN Master — not found</div>`;
          logStep_(`Source: ASIN Master not found (${asin})`);
        }
      }
      body.appendChild(s1el);
      addCollapseListeners_(s1el);

      // Sheet2
      const s2el = document.createElement('div');
      if (allSources === null) {
        s2el.innerHTML = `<div style="font-size:11px;color:#c00;padding:2px 0;">Sheet2 — fetch error</div>`;
        logStep_(`Source: Sheet2 fetch error (${asin})`);
      } else {
        const s2 = allSources.sheet2;
        if (s2) {
          s2el.innerHTML = buildSourceBlock_('✓ Sheet2', SHEET2_LINK, s2);
          logStep_(`Source: Sheet2 found (${asin})`);
        } else {
          s2el.innerHTML = `<div style="font-size:11px;color:#bbb;padding:2px 0;">✗ Sheet2 — not found</div>`;
          logStep_(`Source: Sheet2 not found (${asin})`);
        }
      }
      body.appendChild(s2el);
      addCollapseListeners_(s2el);

      // Amazon (async — reuse if already fetched during Product Info lookup)
      const amzEl = document.createElement('div');
      amzEl.innerHTML = `<div style="font-size:11px;color:#aaa;padding:2px 0;">Amazon — checking…</div>`;
      body.appendChild(amzEl);

      function setAmz(product, url) {
        if (product) lastAmazonProduct = product; // fallback for 대분류/생산업체/원산지정보
        if (!amzEl.isConnected) return;
        if (product) {
          amzEl.innerHTML = buildSourceBlock_('✓ Amazon', url || null, product);
          logStep_(`Source: Amazon found (${asin})`);
        } else {
          amzEl.innerHTML = `<div style="font-size:11px;color:#bbb;padding:2px 0;">✗ Amazon — not found</div>`;
          logStep_(`Source: Amazon not found (${asin})`);
        }
        addCollapseListeners_(amzEl);
      }

      // amazonProduct is undefined → not yet fetched (sheet1/sheet2 path); null → fetched+not found
      if (r.amazonProduct !== undefined) {
        setAmz(r.amazonProduct, r.amazonUrl);
      } else {
        logStep_(`Source: Amazon fetching… (${asin})`);
        fetchAmazonProduct_(asin, setAmz);
      }
    });
  }

  // ── Styles ───────────────────────────────────────────────────────────────
  // GM_addStyle may be undefined if Chrome MV3 restricts Tampermonkey grants;
  // fall back to a plain <style> element so the UI still renders.
  function safeAddStyle_(css) {
    try { if (typeof GM_addStyle === 'function') { GM_addStyle(css); return; } } catch (_) {}
    try {
      const el = document.createElement('style');
      el.textContent = css;
      (document.head || document.documentElement || document.body).appendChild(el);
    } catch (_) {}
  }
  safeAddStyle_(`
    /* ── Main panel ─────────────────────────────────────────────────────── */
    #sp-order-panel {
      position: fixed;
      right: 16px;
      top: 72px;
      width: 330px;
      min-width: 280px;
      max-width: 700px;
      max-height: calc(100vh - 80px);
      display: flex;
      flex-direction: column;
      overflow: hidden;
      background: rgb(244,246,251);
      border: 1px solid rgba(0,0,0,0.10);
      border-radius: 22px;
      box-shadow:
        0 10px 28px rgba(0,0,0,0.10),
        0 3px 8px rgba(0,0,0,0.06),
        inset 0 1px 0 rgba(255,255,255,0.70);
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
      font-size: 12.5px;
      color: #1c1c1e;
      z-index: 99999;
    }

    #sp-order-panel * { box-sizing: border-box; }

    /* ── Header ─────────────────────────────────────────────────────────── */
    #sp-panel-header {
      position: relative;
      z-index: 1;
      padding: 9px 12px;
      background: rgb(236,239,248);
      border-bottom: 1px solid rgba(255, 255, 255, 0.22);
      border-radius: 22px 22px 0 0;
      display: flex;
      align-items: center;
      gap: 7px;
      cursor: move;
      font-weight: 600;
      font-size: 13px;
      user-select: none;
    }
    #sp-abm-log-btn {
      margin-left: auto;
      cursor: pointer;
      opacity: .5;
      font-size: 10px;
      font-weight: 700;
      letter-spacing: 0.4px;
      line-height: 1.4;
      padding: 3px 5px;
      border-radius: 6px;
    }
    #sp-abm-log-btn:hover { opacity: 1; background: rgba(180,185,210,0.35); }
    #sp-abm-log-btn.sp-abm-has-pending { opacity: 1; color: #c0392b; }
    #sp-settings-btn {
      cursor: pointer;
      opacity: .5;
      font-size: 14px;
      line-height: 1.4;
      padding: 2px 5px;
      border-radius: 6px;
    }
    #sp-settings-btn:hover { opacity: 1; background: rgba(180,185,210,0.35); }
    #sp-minimize-btn {
      cursor: pointer;
      opacity: .5;
      font-size: 18px;
      line-height: .7;
      padding: 2px 6px 4px;
      border-radius: 6px;
    }
    #sp-minimize-btn:hover { opacity: 1; background: rgba(180,185,210,0.35); }
    #sp-panel-close {
      cursor: pointer;
      opacity: .5;
      font-size: 15px;
      line-height: 1;
      padding: 2px 5px;
      border-radius: 6px;
    }
    #sp-panel-close:hover { opacity: 1; background: rgba(180,185,210,0.35); }

    /* ── Settings drawer ──────────────────────────────────────────────────── */
    #sp-settings-drawer {
      max-height: 0;
      overflow: hidden;
      transition: max-height 0.25s ease;
      position: relative;
      z-index: 2;
    }
    #sp-settings-drawer.sp-settings-open {
      max-height: 220px;
      border-bottom: 1px solid rgba(255,255,255,0.18);
    }
    .sp-settings-inner {
      padding: 8px 12px 10px;
      background: rgba(255,255,255,0.06);
    }
    /* Body text is selectable/copyable (drag the title bar to move the panel). */
    #sp-panel-body, #sp-panel-body .sp-val, #sp-panel-body .sp-label,
    #sp-panel-body .sp-item-block, #sp-panel-body #sp-notes-content {
      -webkit-user-select: text;
      user-select: text;
      cursor: auto;
    }
    /* Collapsible section headers stay click-to-toggle, not selectable. */
    #sp-panel-body .sp-block-title { -webkit-user-select: none; user-select: none; cursor: pointer; }

    #sp-order-panel.minimized #sp-panel-body { display: none; }
    #sp-order-panel.minimized #sp-settings-drawer { display: none; }
    #sp-order-panel.minimized:not(.sp-docked) #sp-panel-header { border-radius: 22px; border-bottom: none; cursor: pointer; }

    /* ── Edge-resize grips (invisible overlay divs, bypass scrollbar) ────── */
    .sp-re { position: absolute; z-index: 9998; pointer-events: auto; }
    .sp-re-n  { top: 0;    left: 10px;  right: 10px;  height: 8px; cursor: n-resize; }
    .sp-re-s  { bottom: 0; left: 10px;  right: 10px;  height: 8px; cursor: s-resize; }
    .sp-re-e  { right: 0;  top: 10px;   bottom: 10px; width:  8px; cursor: e-resize; }
    .sp-re-w  { left: 0;   top: 10px;   bottom: 10px; width:  8px; cursor: w-resize; }
    .sp-re-ne { top: 0;    right: 0;    width: 14px;  height: 14px; cursor: ne-resize; }
    .sp-re-nw { top: 0;    left: 0;     width: 14px;  height: 14px; cursor: nw-resize; }
    .sp-re-se { bottom: 0; right: 0;    width: 14px;  height: 14px; cursor: se-resize; }
    .sp-re-sw { bottom: 0; left: 0;     width: 14px;  height: 14px; cursor: sw-resize; }
    /* Buttons must sit above edge grips so clicks always reach them */
    #sp-abm-log-btn, #sp-settings-btn, #sp-minimize-btn, #sp-panel-close { position: relative; z-index: 10000; }
    /* Hide lower grips when minimized (body is hidden, bottom edge is gone) */
    #sp-order-panel.minimized .sp-re-s,
    #sp-order-panel.minimized .sp-re-se,
    #sp-order-panel.minimized .sp-re-sw,
    #sp-order-panel.minimized .sp-re-e,
    #sp-order-panel.minimized .sp-re-w  { display: none; }

    /* ── Scrollable body ────────────────────────────────────────────────── */
    #sp-panel-body {
      position: relative;
      z-index: 1;
      padding: 10px 14px 8px;
      overflow-y: auto;
      flex: 1;
      min-height: 0;
    }

    #sp-order-panel.sp-compact .sp-row { flex-direction: column; gap: 1px; }
    #sp-order-panel.sp-compact .sp-label { min-width: 0; font-size: 10.5px; }
    #sp-order-panel.sp-compact .sp-val   { font-size: 11.5px; }

    /* ── Input bars ─────────────────────────────────────────────────────── */
    #sp-id-bar {
      display: flex;
      gap: 6px;
      align-items: center;
      margin-bottom: 8px;
    }
    #sp-order-input {
      flex: 1;
      background: rgba(255,255,255,0.38);
      border: 1px solid rgba(180,190,220,0.40);
      border-radius: 8px;
      padding: 5px 8px;
      font-size: 12px;
      font-family: monospace;
      outline: none;
      color: #1a1a2e;
    }
    #sp-order-input:focus { border-color: #5ba4cf; box-shadow: 0 0 0 2px rgba(91,164,207,.20); background: rgba(255,255,255,0.58); }
    #sp-asin-input {
      flex: 1;
      background: rgba(255,255,255,0.38);
      border: 1px solid rgba(180,190,220,0.40);
      border-radius: 8px;
      padding: 5px 8px;
      font-size: 12px;
      font-family: monospace;
      outline: none;
      color: #1a1a2e;
    }
    #sp-asin-input:focus { border-color: #f0a500; box-shadow: 0 0 0 2px rgba(240,165,0,.20); background: rgba(255,255,255,0.58); }

    /* ── Buttons ────────────────────────────────────────────────────────── */
    #sp-lookup-btn {
      background: rgba(91,164,207,0.88);
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 5px 10px;
      cursor: pointer;
      font-size: 12px;
      white-space: nowrap;
      backdrop-filter: blur(4px);
    }
    #sp-lookup-btn:hover { background: rgba(74,143,186,0.95); }
    #sp-product-btn {
      background: rgba(240,165,0,0.88);
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 5px 10px;
      cursor: pointer;
      font-size: 12px;
      white-space: nowrap;
      backdrop-filter: blur(4px);
    }
    #sp-product-btn:hover { background: rgba(217,146,0,0.95); }

    #sp-autofill-bar { margin-bottom: 8px; display: none; }
    #sp-autofill-btn {
      background: rgba(39,174,96,0.88);
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 5px 0;
      cursor: pointer;
      font-size: 12px;
      width: 100%;
      backdrop-filter: blur(4px);
    }
    #sp-autofill-btn:hover:not(:disabled) { background: rgba(33,154,82,0.96); }
    #sp-autofill-btn:disabled { background: rgba(168,213,181,0.7); cursor: default; }
    #sp-fill-status {
      display: none;
      font-size: 11px;
      color: #27ae60;
      margin-top: 4px;
      text-align: center;
    }

    /* ── Order ID chips ─────────────────────────────────────────────────── */
    #sp-detected-ids { margin-bottom: 8px; display: flex; flex-wrap: wrap; gap: 4px; min-height: 0; }
    .sp-chip {
      background: rgba(91,164,207,0.12);
      border: 1px solid rgba(91,164,207,0.30);
      color: #2a6496;
      border-radius: 12px;
      padding: 2px 10px;
      font-size: 11.5px;
      cursor: pointer;
      font-family: monospace;
      user-select: none;
      backdrop-filter: blur(8px);
    }
    .sp-chip:hover { background: rgba(91,164,207,0.22); }

    #sp-status {
      text-align: center;
      padding: 14px 8px;
      color: #6e6e73;
      font-size: 12px;
    }

    /* ── Collapsible blocks ─────────────────────────────────────────────── */
    .sp-block { margin-top: 4px; }
    .sp-block-title {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 7px 0 4px;
      font-weight: 600;
      font-size: 12.5px;
      color: #1c1c1e;
      cursor: pointer;
      user-select: none;
      border-top: 1px solid rgba(0,0,0,0.08);
    }
    .sp-block-title .sp-chevron { margin-left: auto; transition: transform .18s; color: rgba(110,110,115,0.8); }
    .sp-block.collapsed .sp-block-title .sp-chevron { transform: rotate(-90deg); }
    .sp-block.collapsed .sp-block-body { display: none; }

    /* ── Data rows ──────────────────────────────────────────────────────── */
    .sp-row {
      display: flex;
      align-items: flex-start;
      padding: 4px 0;
      gap: 6px;
    }
    .sp-row:nth-child(odd) {
      background: rgba(255,255,255,0.38);
      margin: 0 -14px;
      padding: 4px 14px;
      border-radius: 6px;
    }
    .sp-label { color: #3a6ea8; min-width: 128px; flex-shrink: 0; font-size: 12px; }
    .sp-val   { color: #1c1c1e; font-weight: 500; word-break: break-all; font-size: 12px; }
    .sp-val.link { color: #3a6ea8; text-decoration: underline; cursor: pointer; }

    /* Per-row one-click copy button */
    .sp-copy {
      margin-left: auto;
      align-self: flex-start;
      flex-shrink: 0;
      background: none;
      border: none;
      padding: 0 2px;
      cursor: pointer;
      color: #9aa3b2;
      opacity: 0.4;
      line-height: 1;
      transition: opacity 0.12s, color 0.12s;
    }
    .sp-row:hover .sp-copy { opacity: 0.85; }
    .sp-copy:hover { color: #3a6ea8; opacity: 1; }
    .sp-copy.copied { color: #27ae60; opacity: 1; font-size: 12px; }

    .sp-items-title {
      font-weight: 600;
      font-size: 11.5px;
      color: #6e6e73;
      padding: 6px 0 2px;
      border-top: 1px solid rgba(0,0,0,0.07);
      margin-top: 2px;
    }
    .sp-item-block {
      padding: 4px 0 6px;
      border-bottom: 1px solid rgba(0,0,0,0.05);
    }
    .sp-item-block:last-child { border-bottom: none; }
    .sp-item-title {
      font-size: 12px;
      font-weight: 500;
      color: #1c1c1e;
      line-height: 1.4;
      word-break: break-word;
      padding: 2px 0 4px;
    }
    .sp-item-title a {
      color: #5ba4cf;
      text-decoration: underline;
    }

    /* ── Step log ───────────────────────────────────────────────────────── */
    #sp-load-log {
      flex-shrink: 0;
      position: relative;
      z-index: 2;
    }
    #sp-load-log:not(:has(#sp-log-entries > *)) { display: none; }
    #sp-log-entries {
      font-size: 10px;
      color: #6e6e73;
      padding: 3px 14px 4px;
      border-top: 1px solid rgba(0,0,0,0.07);
      max-height: 56px;
      overflow-y: auto;
      font-family: monospace;
      line-height: 1.6;
    }
    #sp-load-log.sp-log-collapsed { display: none; }

    /* ── Toggle button (shown when panel is closed) ──────────────────────── */
    #sp-toggle-btn {
      position: fixed;
      right: 16px;
      top: 56px;
      background: rgba(91,164,207,0.82);
      backdrop-filter: blur(18px) saturate(160%);
      -webkit-backdrop-filter: blur(18px) saturate(160%);
      color: #fff;
      border: 1px solid rgba(255,255,255,0.4);
      border-radius: 20px;
      padding: 6px 12px;
      font-size: 12px;
      cursor: pointer;
      z-index: 99999;
      box-shadow: 0 2px 12px rgba(0,0,0,.18);
      display: none;
    }
    #sp-toggle-btn:hover { background: rgba(74,143,186,0.92); }

    /* ── Docked mode: render inline inside Zendesk's Apps panel ───────────── */
    #sp-order-panel.sp-docked {
      position: relative !important;
      width: 100% !important;
      max-width: 100% !important;
      height: auto;
      max-height: none !important;
      flex-shrink: 0 !important;
      left: auto !important; top: auto !important; right: auto !important;
      margin: 0 0 10px 0;
      border-radius: 10px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.08);
      animation: none !important;
      z-index: auto;
    }
    /* In docked mode, hide all resize grips except the south (bottom-drag) one */
    #sp-order-panel.sp-docked .sp-re:not(.sp-re-s) { display: none !important; }
    #sp-order-panel.sp-docked .sp-re-s {
      display: block;
      height: 8px;
      cursor: row-resize;
      background: transparent;
      border-radius: 0 0 10px 10px;
      transition: background 0.15s;
    }
    #sp-order-panel.sp-docked .sp-re-s:hover { background: rgba(0,0,0,0.07); }
    #sp-order-panel.sp-docked.minimized .sp-re-s { display: none !important; }
    /* Native Zendesk app-section chrome look (matches ChannelReply's block). */
    #sp-order-panel.sp-docked {
      border: 1px solid #e9ebed;
      box-shadow: none;
      background: #fff !important;
    }
    #sp-order-panel.sp-docked #sp-panel-header {
      cursor: pointer;
      border-radius: 9px 9px 0 0;
      background: #fff;
      color: #2f3941;
      border-bottom: 1px solid #e9ebed;
      font-size: 14px;
      font-weight: 700;
      padding: 11px 14px;
    }
    #sp-order-panel.sp-docked.minimized #sp-panel-header { border-radius: 9px; }
    #sp-order-panel.sp-docked #sp-panel-header svg:first-of-type { width: 16px; height: 16px; }
    /* Hide floating-only controls when docked */
    #sp-order-panel.sp-docked #sp-minimize-btn,
    #sp-order-panel.sp-docked #sp-panel-close { display: none !important; }
    #sp-order-panel.sp-docked #sp-abm-log-btn { margin-left: auto; }
    /* Settings drawer: override colors for the white docked background */
    #sp-order-panel.sp-docked #sp-settings-drawer.sp-settings-open {
      border-bottom: 1px solid #e9ebed;
      max-height: 280px;
    }
    #sp-order-panel.sp-docked .sp-settings-inner { background: #f8f9fa; }
    /* Panel body scrolls within the docked panel (same as floating mode).
       The panel itself is capped at max-height above, so overflow activates. */
    #sp-order-panel.sp-docked #sp-panel-body { overflow-y: auto !important; max-height: none !important; flex: 1 !important; min-height: 0 !important; background: #fff; }
    /* Collapse/expand toggle button (only visible when docked) */
    #sp-dock-collapse-btn {
      display: none;
      align-items: center;
      justify-content: center;
      width: 30px;
      height: 30px;
      background: transparent;
      border: 1px solid transparent;
      border-radius: 4px;
      cursor: pointer;
      color: #68737d;
      padding: 0;
      flex-shrink: 0;
      position: relative;
      z-index: 10000;
    }
    #sp-dock-collapse-btn:hover { background: rgba(0,0,0,0.05); border-color: #d8dcde; }
    #sp-dock-collapse-btn svg { transform: rotate(180deg); transition: transform 0.18s; }
    #sp-order-panel.sp-docked #sp-dock-collapse-btn { display: inline-flex; }
    #sp-order-panel.sp-docked.minimized #sp-dock-collapse-btn svg { transform: rotate(0deg); }

    #sp-mcf-bar { margin-bottom: 8px; display: none; }
    #sp-mcf-btn {
      background: rgba(255,153,0,0.88);
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 5px 0;
      cursor: pointer;
      font-size: 12px;
      width: 100%;
      backdrop-filter: blur(4px);
    }
    #sp-mcf-btn:hover { background: rgba(230,138,0,0.96); }
    #sp-mcf-status {
      display: none;
      font-size: 11px;
      color: #27ae60;
      margin-top: 4px;
      text-align: center;
    }

    #sp-nrn-bar { margin-top: 6px; }
    #sp-nrn-btn {
      background: rgba(108,117,140,0.90);
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 5px 0;
      cursor: pointer;
      font-size: 12px;
      width: 100%;
      backdrop-filter: blur(4px);
    }
    #sp-nrn-btn:hover:not(:disabled) { background: rgba(86,94,115,0.96); }
    #sp-nrn-btn:disabled { background: rgba(160,165,175,0.55); color: rgba(255,255,255,0.8); cursor: not-allowed; }
    #sp-nrn-status {
      display: none;
      font-size: 11px;
      color: #27ae60;
      margin-top: 4px;
      text-align: center;
    }

    #sp-ai-reason-bar { margin-top: 6px; border-top: 1px solid #e9ebec; padding-top: 7px; }
    #sp-ai-reason-btn {
      background: rgba(124,58,237,0.88);
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 5px 0;
      cursor: pointer;
      font-size: 12px;
      width: 100%;
      backdrop-filter: blur(4px);
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 5px;
    }
    #sp-ai-reason-btn:hover:not(:disabled) { background: rgba(100,40,200,0.96); }
    #sp-ai-reason-btn:disabled { background: rgba(160,120,230,0.65); cursor: default; }
    .sp-ai-star { display: inline-block; }
    @keyframes sp-spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    .sp-ai-spin { animation: sp-spin 1s linear infinite; }
    .sp-combo-item:hover { background: #eef7f1; }

    #sp-notes-bar {
      margin-bottom: 6px;
    }
    #sp-notes-bar label {
      display: flex;
      align-items: center;
      gap: 5px;
      cursor: pointer;
      user-select: none;
      font-size: 12px;
      color: #4a8fcf;
      font-weight: 500;
    }
    #sp-notes-toggle, #sp-seller-notes-toggle { cursor: pointer; accent-color: #5ba4cf; }
    #sp-notes-section, #sp-seller-notes-section {
      display: none;
      margin-bottom: 8px;
    }
    #sp-notes-content, #sp-seller-notes-content {
      font-size: 12px;
      color: #1c1c1e;
      white-space: pre-wrap;
      padding: 6px 8px;
      background: rgba(255,255,255,0.40);
      border: 1px solid rgba(0,0,0,0.08);
      border-radius: 8px;
      min-height: 36px;
      width: 100%;
      box-sizing: border-box;
      font-family: inherit;
      resize: vertical;
      cursor: default;
    }

    #sp-order-panel.sp-dragging { box-shadow: 0 8px 20px rgba(0,0,0,0.10); }
  `);

  // ── Raw ABM comment collapse (see collapseRawAbmComments_ below) ──────────
  // Never touches the Zendesk comment article's own children — only toggles
  // this class on the article itself (safe className mutation, no structural
  // change) plus a standalone sibling toggle button.
  //
  // display:none, not max-height/opacity (tried first, reverted — see v3.5.4):
  // the article is a CSS grid container that Zendesk itself sets
  // `min-height: max-content` on. Per the CSS2.1 box-height algorithm,
  // min-height wins whenever it conflicts with max-height, so `max-height: 0`
  // was silently ignored — confirmed live via getComputedStyle on ticket
  // #1000154702: maxHeight computed as "0px" but the actual layout height
  // stayed at ~1195px. The article turned invisible (opacity:0) but kept
  // occupying its full height, leaving a large blank gap instead of
  // collapsing. display:none isn't subject to that min/max-height
  // precedence rule at all, so it collapses reliably regardless of Zendesk's
  // own grid/min-height styling on the element.
  safeAddStyle_(`
    article.gcx-abm-raw-collapsed {
      display: none !important;
    }
    .gcx-abm-raw-toggle {
      font-size: 11px;
      color: #68737d;
      cursor: pointer;
      padding: 2px 0 6px 52px;
      user-select: none;
    }
    .gcx-abm-raw-toggle:hover { text-decoration: underline; }
  `);

  // ── Panel HTML ────────────────────────────────────────────────────────────
  function buildPanel() {
    const d = document.createElement('div');
    d.id = PANEL_ID;
    d.innerHTML = `
      <div class="sp-re sp-re-n"></div><div class="sp-re sp-re-ne"></div>
      <div class="sp-re sp-re-e"></div><div class="sp-re sp-re-se"></div>
      <div class="sp-re sp-re-s"></div><div class="sp-re sp-re-sw"></div>
      <div class="sp-re sp-re-w"></div><div class="sp-re sp-re-nw"></div>
      <div id="sp-panel-header">
        <svg width="18" height="18" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
          <text x="3" y="38" font-size="38" font-family="Georgia,serif" font-style="italic" fill="#FF9900">a</text>
          <path d="M6 40 Q24 48 42 40" stroke="#FF9900" stroke-width="3" fill="none" stroke-linecap="round"/>
        </svg>
        GCX Reply
        <span style="font-size:9.5px;font-weight:normal;color:#bbb;margin-left:3px;vertical-align:middle;letter-spacing:0.3px;">v${SCRIPT_VER}</span>
        <span id="sp-abm-log-btn" title="ABM relay log">ABM</span>
        <span id="sp-settings-btn" title="Settings">⚙</span>
        <span id="sp-minimize-btn" title="Minimize">─</span>
        <span id="sp-panel-close" title="Close">✕</span>
        <button id="sp-dock-collapse-btn" type="button" aria-label="Collapse section" data-garden-id="buttons.icon_button" data-garden-version="9.14.2">
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16" aria-hidden="true" focusable="false" data-garden-id="buttons.icon" data-garden-version="9.14.2">
            <path fill="currentColor" d="M12.688 5.61a.5.5 0 0 1 .69.718l-.066.062-5 4a.5.5 0 0 1-.542.054l-.082-.054-5-4a.5.5 0 0 1 .55-.83l.074.05L8 9.359l4.688-3.75z"/>
          </svg>
        </button>
      </div>
      <div id="sp-settings-drawer"></div>
      <div id="sp-panel-body">
        <div id="sp-id-bar">
          <input id="sp-order-input" type="text" placeholder="408-XXXXXXX-XXXXXXX" maxlength="19"/>
          <button id="sp-lookup-btn">Lookup</button>
        </div>
        <div id="sp-id-bar" style="margin-bottom:10px;">
          <input id="sp-asin-input" type="text" placeholder="ASIN (B0XXXXXXXXX)" maxlength="10"/>
          <button id="sp-product-btn">Product</button>
        </div>
        <div id="sp-detected-ids"></div>
        <div id="sp-autofill-bar">
          <button id="sp-autofill-btn">Auto-Fill Form</button>
          <div id="sp-fill-status"></div>
        </div>
        <div id="sp-mcf-bar">
          <button id="sp-mcf-btn">→ MCF</button>
          <div id="sp-mcf-status"></div>
        </div>
        <div id="sp-nrn-bar">
          <button id="sp-nrn-btn" disabled title="Only available on Amazon Buyer Message tickets">Mark as NRN</button>
          <div id="sp-nrn-status"></div>
        </div>
        <div id="sp-notes-section">
          <textarea id="sp-notes-content" placeholder="(no notes)" readonly></textarea>
        </div>
        <div id="sp-seller-notes-section">
          <div id="sp-seller-notes-wrap"><div class="sp-row"><span class="sp-val">${SPINNER_HTML}</span></div></div>
        </div>
        <div id="sp-ai-reason-result"></div>
        <div id="sp-result">
          <div id="sp-status">Scanning ticket for order IDs…</div>
        </div>
        <div id="sp-product-result"></div>
      </div>
      <div id="sp-load-log"><div id="sp-log-entries"></div></div>
    `;
    return d;
  }

  // ── Format helpers ────────────────────────────────────────────────────────
  function esc(s) {
    if (!s) return '';
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  function fmtDate(iso) {
    if (!iso) return '—';
    return new Date(iso).toLocaleString('sv-SE', { timeZone: 'UTC' }).slice(0, 16).replace('T', ' ');
  }

  // Seller Central shows purchase dates in the marketplace's local timezone, not UTC.
  // SP-API returns UTC. Convert so GCX Reply matches Seller Central.
  const PURCHASE_TZ_ = {
    IN: 'Asia/Kolkata',     // IST = UTC+5:30
    JP: 'Asia/Tokyo',       // JST = UTC+9
    SG: 'Asia/Singapore',   // SGT = UTC+8
    AU: 'Australia/Sydney', // AEST/AEDT
    KR: 'Asia/Seoul',       // KST = UTC+9
  };
  function purchaseDateLocal_(isoUtc, countryCode) {
    if (!isoUtc) return '';
    const tz = PURCHASE_TZ_[countryCode];
    if (!tz) return isoUtc.slice(0, 10);
    return new Date(isoUtc).toLocaleDateString('sv-SE', { timeZone: tz });
  }
  function fmtPurchaseDate_(isoUtc, countryCode) {
    if (!isoUtc) return '—';
    const tz = PURCHASE_TZ_[countryCode];
    if (!tz) return fmtDate(isoUtc) + ' (UTC)';
    const local = new Date(isoUtc).toLocaleString('sv-SE', { timeZone: tz }).slice(0, 16).replace('T', ' ');
    const label = { IN:'IST', JP:'JST', SG:'SGT', AU:'AEST', KR:'KST' }[countryCode] || tz;
    return `${local} (${label})`;
  }

  function fmtShipRange(earliest, latest) {
    if (!earliest) return '—';
    const fmt = iso => new Date(iso).toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' });
    const e = fmt(earliest), l = latest ? fmt(latest) : '';
    return (!l || e === l) ? e : `${e} – ${l}`;
  }

  // Renders in the browser's local timezone (matches ChannelReply's own display,
  // which does the same client-side conversion rather than a fixed timezone).
  function fmtCrDateTime_(iso) {
    if (!iso) return null;
    const d = new Date(iso);
    if (isNaN(d.getTime())) return iso;
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }

  function row(label, value, isLink) {
    const has = value != null && value !== '' && value !== '—';
    return `<div class="sp-row">
      <span class="sp-label">${esc(label)}</span>
      <span class="sp-val${isLink ? ' link' : ''}">${esc(value) || '—'}</span>
      ${has ? COPY_BTN : ''}
    </div>`;
  }

  function fulfillmentLabel_(channel, sku) {
    if (sku && /pan|eup/i.test(sku)) return 'PAN EU';
    if (channel === 'AFN') return 'FBA';
    if (channel === 'MFN') return 'Merchant (FBM)';
    return channel || '—';
  }

  function rowReturnAsin(asinStr, salesChannel, itemsStatus) {
    if (!asinStr || asinStr === '—') {
      const note = itemsStatus === 403
        ? `<span style="font-size:10.5px;color:#e67e22;margin-left:4px;">(GetOrderItems 권한 필요)</span>`
        : '';
      return `<div class="sp-row"><span class="sp-label">Return ASIN</span><span class="sp-val">—${note}</span></div>`;
    }
    const ch = (salesChannel || 'amazon.com').toLowerCase();
    const links = asinStr.split(',').map(a => a.trim()).filter(Boolean).map(asin => {
      const url = `https://www.${ch}/dp/${asin}`;
      return `<a href="${url}" target="_blank" rel="noopener" style="color:#5ba4cf;text-decoration:underline;">${esc(asin)}</a>`;
    }).join(', ');
    return `<div class="sp-row"><span class="sp-label">Return ASIN</span><span class="sp-val">${links}</span>${COPY_BTN}</div>`;
  }

  function rowLinked(label, text, url) {
    const has = text != null && text !== '' && text !== '—';
    const cell = url
      ? `<a href="${url}" target="_blank" rel="noopener" style="color:#5ba4cf;text-decoration:underline;font-weight:500;">${esc(text)}</a>`
      : `<span class="sp-val">${esc(text) || '—'}</span>`;
    return `<div class="sp-row"><span class="sp-label">${esc(label)}</span>${cell}${has ? COPY_BTN : ''}</div>`;
  }

  // Render selling-marketplace badges from market spreadsheet check
  function marketplacesRow_(mkts) {
    if (!mkts || !mkts.length) {
      return `<div class="sp-row"><span class="sp-label">판매 마켓</span><span class="sp-val" style="color:#aaa;">—</span></div>`;
    }
    const SS_ID  = '172fDVw4tu-hgbpV5FShWj4_SAMxeB54-v5BUlVgJUoA';
    const badges = mkts.map(m => {
      const name = typeof m === 'string' ? m : m.name;
      const gid  = typeof m === 'string' ? null : m.gid;
      const cell = typeof m === 'object' ? m.cell : null;
      let url = `https://docs.google.com/spreadsheets/d/${SS_ID}/edit`;
      if (gid != null) url += `#gid=${gid}`;
      if (cell)        url += `&range=${cell}`;
      return `<a href="${esc(url)}" target="_blank" rel="noopener"
        style="display:inline-block;background:#27ae60;color:#fff;font-size:10px;padding:1px 6px;border-radius:3px;margin-right:3px;margin-bottom:2px;text-decoration:none;">${esc(name)}</a>`;
    }).join('');
    return `<div class="sp-row"><span class="sp-label">판매 마켓</span><span class="sp-val">${badges}</span></div>`;
  }

  // ── Render order data ─────────────────────────────────────────────────────
  function renderOrder(data, orderId, panelAsin, options) {
    options = options || {};
    const showOrder = options.showOrder !== false;
    const showShipping = options.showShipping !== false;
    const prefs = getDataFetchPrefs();

    const o  = data.order   || {};
    const it = data.items   || [];
    const ad = data.address || {};
    const b  = data.buyer   || {};
    const buyerEmail   = b.BuyerEmail || '';
    const scSearchUrl  = sellerCentralSearchUrl_(o.SalesChannel, ad.CountryCode, buyerEmail);

    const itemAsins    = it.map(i => i.ASIN).filter(Boolean);
    const returnAsin   = itemAsins.length ? itemAsins.join(', ') : (panelAsin || '—');
    // All Seller SKUs (multi-item orders list every SKU, deduped, in item order).
    const sellerSkus   = [...new Set(it.map(i => i.SellerSKU).filter(Boolean))];
    const sellerSkuStr = sellerSkus.join(', ');
    // PAN-EU detection: treat the order as PAN if ANY item carries a PAN/EUP SKU.
    const fulfillLabel = fulfillmentLabel_(o.FulfillmentChannel, sellerSkus.join(' ') || it[0]?.SellerSKU || '');
    const amount     = o.OrderTotal ? `${o.OrderTotal.Amount} ${o.OrderTotal.CurrencyCode}` : '—';
    const buyerName  = b.BuyerName || o.BuyerInfo?.BuyerName || ad.Name || getTicketRequesterNameFromDom_() || '—';

    const addrParts = [ad.Name, ad.AddressLine1, ad.AddressLine2, ad.AddressLine3,
                       [ad.City, ad.StateOrRegion, ad.PostalCode].filter(Boolean).join(' '),
                       ad.CountryCode].filter(Boolean);

    const addrRows = addrParts.map(p =>
      `<div class="sp-row"><span class="sp-val">${esc(p)}</span></div>`
    ).join('');

    const itemRows = it.map(item => {
      const asin  = item.ASIN || '';
      const title = item.Title || asin;
      const url   = asin ? `https://www.${(o.SalesChannel || 'amazon.com').toLowerCase()}/dp/${asin}` : '';
      return `<div class="sp-item-block">
        <div class="sp-item-title">${url ? `<a href="${esc(url)}" target="_blank" rel="noopener">${esc(title)}</a>` : esc(title)}</div>
        ${row('Amazon Order Item ID',     item.AmazonOrderItemId || item.OrderItemId)}
        ${row('ASIN',                     asin)}
        ${row('Seller SKU',               item.SellerSKU)}
        ${row('Quantity Ordered',         item.QuantityOrdered)}
        ${row('Item Price',               item.ItemPrice)}
        ${row('Item Price With Discount', item.ItemPriceWithDiscount)}
        ${row('Item Tax',                 item.ItemTax)}
      </div>`;
    }).join('');

    // Prefer total ITEM quantity across orders (matches the '✅전체 주문'/'❎전체 환불'
    // ZD field values — see fetchScBuyerStats_/checkRefunds_) over raw order count —
    // a customer with 4 orders totaling 5 items should show 5, not 4. Previously this
    // always showed the raw order count, silently diverging from what Auto-Fill used.
    const dispPurchases = data.totalItemsPurchased ?? data.totalPurchases;
    const dispRefunds    = data.totalItemsRefunded  ?? data.totalRefunds;
    const orderCountNote = dispPurchases != null
      ? ` <span id="sp-stat-badge" style="color:#888;font-size:11px;">(구매 ${dispPurchases}건 / 환불 ${dispRefunds}건)</span>`
      : data.orderCount != null
        ? ` <span id="sp-stat-badge" style="color:#888;font-size:11px;">(총 ${data.orderCount}건)</span>`
        : ` <span id="sp-stat-badge" style="color:#888;font-size:11px;">${SPINNER_HTML}</span>`;

    if (!prefs.fetchOrder) {
      return `${rowReturnAsin(returnAsin, o.SalesChannel, data.itemsStatus)}`;
    }

    const shippingSection = !prefs.fetchShipping ? '' : `
          <div class="sp-block collapsed" data-sp-section="shipping">
            <div class="sp-block-title" style="font-size:12px;">
              Shipping Address
              <span class="sp-chevron">▾</span>
            </div>
            <div class="sp-block-body">
              ${row('Amazon Fulfillment Methods', fulfillLabel)}
              ${addrRows || '<div class="sp-row"><span class="sp-val">—</span></div>'}
              ${row('Carrier',               o.Carrier)}
              ${row('Tracking ID',           o.TrackingID)}
              ${row('Fulfillment Center ID', o.FulfillmentCenterId)}
              ${row('Ship Service Level',    o.ShipServiceLevel)}
              ${row('Shipment ID',           o.ShipmentId)}
              ${row('Shipment Item ID',      o.ShipmentItemId)}
              ${row('Shipment Date',         fmtCrDateTime_(o.ShipmentDate))}
              ${row('Buyer Name',            buyerName)}
              ${row('Language',              o.Language)}
            </div>
          </div>`;

    return `
      ${rowReturnAsin(returnAsin, o.SalesChannel, data.itemsStatus)}

      <div class="sp-block" data-sp-section="order">
        <div class="sp-block-title">
          <svg width="16" height="16" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
            <text x="3" y="38" font-size="38" font-family="Georgia,serif" font-style="italic" fill="#FF9900">a</text>
            <path d="M6 40 Q24 47 42 40" stroke="#FF9900" stroke-width="3" fill="none" stroke-linecap="round"/>
          </svg>
          Order${orderCountNote}
          <span class="sp-chevron">▾</span>
        </div>
        <div class="sp-block-body">
          ${rowLinked('Amazon Order ID', orderId, sellerCentralUrl(orderId, o.SalesChannel, ad.CountryCode))}
          ${row('Amazon Order Item ID', o.AmazonOrderItemId)}
          ${sellerSkuStr
              ? row('Seller SKU', sellerSkuStr)
              : `<div class="sp-row" id="sp-seller-sku-row"><span class="sp-label">Seller SKU</span><span class="sp-val">${SPINNER_HTML}</span></div>`}
          ${row('Order Status',     o.OrderStatus)}
          ${row('Purchase Date',    fmtPurchaseDate_(o.PurchaseDate, ad.CountryCode))}
          ${row('Amount',           amount)}
          ${row('Delivery Level',   o.ShipmentServiceLevelCategory || o.ShipServiceLevelCategory)}
          ${row('Ship Date',        fmtShipRange(o.EarliestShipDate, o.LatestShipDate))}
          ${row('Estimated Arrival Date', o.EstimatedArrivalDate)}
          ${row('Sender Email',           o.SenderEmail)}

          ${shippingSection}
          <div id="sp-stat-link-wrap">${
            dispPurchases != null
              ? rowLinked('구매이력 (2yr)', `구매 ${dispPurchases}건 / 환불 ${dispRefunds}건`, scSearchUrl)
              : `<div class="sp-row"><span class="sp-label">구매이력 (2yr)</span><span class="sp-val">${SPINNER_HTML}</span></div>`
          }</div>
          ${it.length > 0 ? `<div class="sp-items-title">Items (${it.length})</div>${itemRows}` : ''}
        </div>
      </div>
    `;
  }

  // ── Seller Central buyer purchase stats (SC session fallback) ────────────
  // SP-API often lacks BuyerEmail PII permission → use the user's logged-in SC
  // session. Hits /orders-api/order/{id} for buyer email, then
  // /orders-api/search?qt=email&q={email} for a 2-year order count.
  function fetchScBuyerStats_(orderId, salesChannel, countryCode, cb) {
    // Seller notes ("For your records only, will not be displayed to the
    // buyer" textarea on the SC order page) — a plain field on this same
    // order response, same as RefundApplied. Threaded through every cb()
    // call below since it's specific to THIS order, not the buyer-wide
    // purchase stats countByEmail_/checkRefunds_ compute.
    // undefined = not fetched yet (or fetch failed); '' = fetched, genuinely
    // empty. Must stay distinguishable — see the extraction line below.
    // Declared FIRST, before any early return: the cache-hit branch below
    // calls countByEmail_/checkRefunds_, whose closures reference this
    // variable — if it were declared after that branch, a cached lookup
    // would skip its initializer entirely and hit the TDZ the moment
    // countByEmail_/checkRefunds_ read it (`Cannot access 'sellerNotes'
    // before initialization`). Same trap class as v3.6.12's
    // MCF_PERSON_OPTIONS bug — see gcx_reply_tampermonkey memory.
    let sellerNotes;

    const scUrl = sellerCentralUrl(orderId, salesChannel, countryCode);
    if (!scUrl) { cb(null); return; }
    const base = scUrl.match(/^https:\/\/[^/]+/)[0];

    // Cache hit: skip the /orders-api/order/{id} call entirely
    if (_scEmailCache_[orderId]) { countByEmail_(_scEmailCache_[orderId]); return; }

    scRequest_({
      method: 'GET',
      url: `${base}/orders-api/order/${orderId}`,
      redirect: 'follow',
      timeout: 15000,
      onload(res) {
        if (res.finalUrl?.includes('/ap/signin') || (res.responseText || '').trimStart().startsWith('<')) {
          showScSessionWarning_(base);
          cb(null);
          return;
        }
        let email = null;
        if (res.status === 200) {
          try {
            const d = JSON.parse(res.responseText);
            email = d.order?.buyerEmail || d.order?.buyer?.email || null;
            // Cache this order's RefundApplied status while we have the response
            const refApplied = d.order?.orderStatus?.RefundApplied;
            if (typeof refApplied === 'boolean') _scRefundCache_[orderId] = refApplied;
            // Cache the item quantity too — caching ONLY RefundApplied here
            // made checkRefunds_ treat this order as fully known and read its
            // quantity as `_scOrderQtyCache_[oid] || 0` → 0, silently dropping
            // the CURRENT order's items from 전체 주문/전체 환불. Confirmed live
            // on ticket #1000154333 (UK, buyer's only order = the current one,
            // 2 items → showed 구매 0건 / q0) and it also explains ticket
            // #1000154377's q3 (4 orders / 5 items, current order's 2 dropped).
            const itemsArr = d.order?.orderItems;
            if (Array.isArray(itemsArr) && itemsArr.length) {
              _scOrderQtyCache_[orderId] = itemsArr.reduce((s, it) => s + (it.QuantityOrdered || 0), 0);
            }
            // ?? not || — an order with genuinely no seller notes returns ""
            // (falsy), and || would collapse that into the same "unknown /
            // not fetched yet" state as undefined, permanently stuck showing
            // the disabled/loading cursor even though the fetch succeeded.
            sellerNotes = d.order?.sellerNotes ?? '';
          } catch {}
          if (!email) {
            const m = res.responseText.match(/"([^"@]+@marketplace\.amazon\.[^"]+)"/);
            if (m) email = m[1];
          }
        }
        if (!email) { cb({ email: null, sellerNotes }); return; }
        _scEmailCache_[orderId] = email;
        countByEmail_(email);
      },
      onerror() { cb(null); },
      ontimeout() { cb(null); },
    });

    function countByEmail_(email) {
      const cached = _scOrdersCache_[email];
      if (cached && Date.now() - cached.ts < 15 * 60 * 1000) {
        checkRefunds_(email, cached.count, cached.orderIds);
        return;
      }
      const now = Date.now();
      const twoYearsAgo = Math.round(now - 2 * 365.25 * 24 * 3600 * 1000);
      GM_xmlhttpRequest({
        method: 'GET',
        url: `${base}/orders-api/search?qt=email&q=${encodeURIComponent(email)}&date-range=${twoYearsAgo}-${now}`,
        redirect: 'follow',
        timeout: 15000,
        onload(res) {
          let count = null;
          let orderIds = [];
          if (res.status === 200) {
            try {
              const d = JSON.parse(res.responseText);
              count = d.total ?? d.totalCount ?? d.totalOrders
                   ?? (Array.isArray(d.orders) ? d.orders.length : null);
              if (Array.isArray(d.orders))
                orderIds = d.orders.slice(0, 50).map(o => o.amazonOrderId).filter(Boolean);
            } catch {
              const m = res.responseText.match(/"(?:total|totalCount|totalOrders)"\s*:\s*(\d+)/);
              if (m) count = parseInt(m[1], 10);
            }
          }
          if (!orderIds.length) { cb({ email, totalPurchases: count, totalRefunds: null, totalItemsPurchased: null, totalItemsRefunded: null, sellerNotes }); return; }
          _scOrdersCache_[email] = { count, orderIds, ts: Date.now() };
          // Fire partial update immediately so purchase count shows while refunds are checked
          cb({ email, totalPurchases: count, totalRefunds: -1, totalItemsPurchased: null, totalItemsRefunded: null, sellerNotes });
          checkRefunds_(email, count, orderIds);
        },
        onerror()  { cb({ email, totalPurchases: null, sellerNotes }); },
        ontimeout() { cb({ email, totalPurchases: null, sellerNotes }); },
      });
    }

    // Per-order refund + item-quantity check: uses _scRefundCache_/_scOrderQtyCache_
    // to skip already-known orders. All uncached orders are fetched in parallel.
    // Item quantity (sum of QuantityOrdered across all orderItems) is used for the
    // '전체 주문'/'전체 환불' ZD fields instead of a raw order count — a customer
    // with 3 orders totaling 8 items should show 8, not 3.
    function checkRefunds_(email, count, orderIds) {
      let refundCount = 0;
      let itemQty = 0;
      let refundItemQty = 0;
      const uncached = [];
      orderIds.forEach(oid => {
        if (_scRefundCache_[oid] !== undefined) {
          const qty = _scOrderQtyCache_[oid] || 0;
          itemQty += qty;
          if (_scRefundCache_[oid]) { refundCount++; refundItemQty += qty; }
        } else {
          uncached.push(oid);
        }
      });
      const finish_ = () => cb({
        email, totalPurchases: count, totalRefunds: refundCount,
        totalItemsPurchased: itemQty, totalItemsRefunded: refundItemQty, sellerNotes,
      });
      if (!uncached.length) { finish_(); return; }

      // Fire in capped-size batches rather than all at once — a customer
      // with a long 2-year order history could otherwise mean 30+ orders'
      // worth of GM_xmlhttpRequest calls hitting Seller Central
      // simultaneously, risking throttling/slow responses. Final aggregated
      // result (itemQty/refundCount/etc) is identical either way; only the
      // in-flight concurrency changes.
      const SC_ORDER_FETCH_CONCURRENCY = 5;
      let idx = 0;
      function runNextBatch_() {
        if (idx >= uncached.length) { finish_(); return; }
        const batch = uncached.slice(idx, idx + SC_ORDER_FETCH_CONCURRENCY);
        idx += batch.length;
        let pending = batch.length;
        const batchDone_ = () => { if (--pending === 0) runNextBatch_(); };
        batch.forEach(oid => {
          scRequest_({
            method: 'GET',
            url: `${base}/orders-api/order/${oid}`,
            redirect: 'follow',
            timeout: 10000,
            onload(r2) {
              try {
                const order = JSON.parse(r2.responseText)?.order || {};
                const val = order.orderStatus?.RefundApplied === true;
                const qty = (order.orderItems || []).reduce((s, it) => s + (it.QuantityOrdered || 0), 0);
                _scRefundCache_[oid] = val;
                _scOrderQtyCache_[oid] = qty;
                itemQty += qty;
                if (val) { refundCount++; refundItemQty += qty; }
              } catch {}
              batchDone_();
            },
            onerror()   { batchDone_(); },
            ontimeout() { batchDone_(); },
          });
        });
      }
      runNextBatch_();
    }
  }

  // Formats an Amazon {Amount, CurrencyCode} money object as "39.99 GBP".
  // Returns null (not '—') for missing input so callers can `||`-fall-back.
  function fmtMoney_(m) {
    return (m && m.Amount != null) ? `${m.Amount} ${m.CurrencyCode || ''}`.trim() : null;
  }

  // ── Seller Central orders-api fallback for ASIN + SKU + price/tax ────────
  // Uses the user's existing SC session cookies — no extra auth needed.
  // Tries marketplace-specific SC domain first, then sellercentral.amazon.com
  // (global SC) as fallback — Spigen accesses most markets via the global domain.
  //
  // Also extracts item price/tax (from each item's ItemCost block) and the
  // order's buyer proxy email — this is SC's own equivalent of what
  // ChannelReply used to provide (ChannelReply is fully cancelled; its API
  // is dead). Confirmed live against real orders that SC's orders-api does
  // NOT expose package/tracking info (Carrier/TrackingID/ShipmentDate/etc.)
  // for FBA orders — `packages` came back null on both a delivered and a
  // still-pending order, since Amazon (not the seller) owns FBA fulfillment/
  // tracking. Those fields stay unpopulated ('—' in the panel); no known
  // ChannelReply-free replacement exists for them.
  function fetchScItems(orderId, salesChannel, countryCode, cb) {
    // scDomain_ already routes EU non-DE → amazon.de, so primaryUrl never hits NL/FR/etc.
    const scPageUrl = sellerCentralUrl(orderId, salesChannel, countryCode);
    const primaryUrl  = scPageUrl ? scPageUrl.replace('/orders-v3/order/', '/orders-api/order/') : null;
    const fallbackUrl = `https://sellercentral.amazon.com/orders-api/order/${orderId}`;

    function parseResult(responseText) {
      const d = JSON.parse(responseText);
      const o = d.order || {};
      const items = (o.orderItems || [])
        .map(it => {
          const c = it.ItemCost || {};
          const price    = c.Subtotal;
          const discount = c.Promotion;
          const withDiscount = (price && discount)
            ? fmtMoney_({ Amount: Math.round((price.Amount - discount.Amount) * 100) / 100, CurrencyCode: price.CurrencyCode })
            : fmtMoney_(price);
          return {
            ASIN:                  it.ASIN,
            SellerSKU:             it.SellerSKU,
            Title:                 it.Title,
            QuantityOrdered:       it.QuantityOrdered,
            OrderItemId:           it.OrderItemId,
            ItemPrice:             fmtMoney_(price),
            ItemPriceWithDiscount: withDiscount,
            ItemTax:               fmtMoney_(c.SubtotalTax || c.Tax),
          };
        })
        .filter(it => it.ASIN);
      const extras = { SenderEmail: o.buyerProxyEmail || null };
      return { items, extras };
    }

    function tryUrl(url, onFail) {
      scRequest_({
        method:  'GET',
        url,
        headers: { 'Accept': 'application/json' },
        redirect: 'follow',
        timeout:  20000,
        onload(res) {
          if (res.finalUrl?.includes('/ap/signin') || (res.status === 200 && (res.responseText || '').trimStart().startsWith('<'))) {
            const base = url.match(/^https:\/\/[^/]+/)?.[0];
            showScSessionWarning_(base);
            logStep_(`SC items: session expired → login redirect`);
            return onFail();
          }
          if (res.status !== 200) {
            logStep_(`SC items: ${url.replace(/^https:\/\//,'')} → HTTP ${res.status}${res.finalUrl && res.finalUrl !== url ? ` (→ ${res.finalUrl.replace(/^https:\/\//,'').slice(0,40)})` : ''}`);
            return onFail();
          }
          try {
            const { items, extras } = parseResult(res.responseText);
            if (items.length) cb(items, extras);
            else {
              const head = (res.responseText || '').trim().slice(0, 60).replace(/\s+/g, ' ');
              logStep_(`SC items: 200 but 0 parsed — body starts: "${head}"`);
              onFail();
            }
          } catch (e) {
            logStep_(`SC items: parse error — ${e.message}`);
            onFail();
          }
        },
        onerror()   { logStep_(`SC items: network error on ${url.replace(/^https:\/\//,'').slice(0,40)}`); onFail(); },
        ontimeout() { logStep_('SC items: timeout'); onFail(); },
      });
    }

    if (primaryUrl && primaryUrl !== fallbackUrl) {
      tryUrl(primaryUrl, () => tryUrl(fallbackUrl, () => cb(null)));
    } else {
      tryUrl(fallbackUrl, () => cb(null));
    }
  }

  // ── Fetch order via GAS ───────────────────────────────────────────────────
  function fetchOrder(orderId, _retries, _fresh) {
    _retries = _retries || 0;
    const _session = _panelSession;
    setStatus('Fetching order data…', true);
    if (!_retries) logStep_(`Fetching order ${orderId}…`);
    if (!_retries) prefetchScOrder_(orderId); // SC order call in parallel (v3.8.0)
    backendRequest_(`orderId=${encodeURIComponent(orderId)}`, {
      timeout:  15000,
      onload(res) {
        if (_panelSession !== _session) return;
        const result = document.getElementById('sp-result');
        if (!result) return;
        if (res.responseText.trimStart().startsWith('<')) {
          if (_retries < 2) {
            logStep_(`GAS not ready — retry ${_retries + 1}/2…`);
            setStatus('Retrying…', true);
            setTimeout(() => fetchOrder(orderId, _retries + 1), 2000);
            return;
          }
          setStatus('GAS error — refresh and try again');
          logStep_('Order fetch: GAS returned error page');
          return;
        }
        try {
          const data = JSON.parse(res.responseText);
          if (data.error) { setStatus(data.error); logStep_('Order error: ' + data.error); return; }

          // Store for auto-fill
          lastOrderData = data;
          learnScBase_(orderId, scBaseFromResult_(orderId, data.order?.SalesChannel, data.address?.CountryCode));
          logStep_(`Order loaded — ${data.order?.SalesChannel || data.region || 'unknown'} | 구매이력: ${data.totalPurchases != null ? `구매 ${data.totalPurchases}건 / 환불 ${data.totalRefunds}건` : 'N/A'}`);
          window.__gcxRefreshNrnState && window.__gcxRefreshNrnState();
          window.__gcxAutoCorrectLanguage && window.__gcxAutoCorrectLanguage();
          maybeShowAutoFill(document.getElementById(PANEL_ID));

          const asinInput = document.getElementById('sp-asin-input');
          const itemAsins = (data.items || []).map(i => i.ASIN).filter(Boolean);

          // Resolve ASIN before rendering: SP-API items > current input > page scan
          let resolvedAsin = asinInput?.value.trim() || '';
          if (itemAsins.length) {
            if (asinInput && !asinInput.value) asinInput.value = itemAsins.join(', ');
            resolvedAsin = asinInput?.value.trim() || itemAsins[0] || '';
          } else if (!resolvedAsin) {
            const pageAsins = [...new Set([...getCachedBodyText_().matchAll(ASIN_RE)].map(m => m[1]))];
            const detected  = pageAsins[0];
            if (detected && asinInput) { asinInput.value = detected; resolvedAsin = detected; }
          }

          result.innerHTML = renderOrder(data, orderId, resolvedAsin);
          result.querySelectorAll('.sp-block-title').forEach(title => {
            title.addEventListener('click', e => {
              e.stopPropagation();
              title.closest('.sp-block').classList.toggle('collapsed');
            });
          });
          applySectionState(result);

          // SC session: always run for refund count; SP-API Finances API may be inaccessible.
          // Uses in-place DOM patch (#sp-stat-badge / #sp-stat-link-wrap) to avoid full re-render.
          fetchScBuyerStats_(orderId, data.order?.SalesChannel, data.address?.CountryCode, stats => {
            if (_panelSession !== _session || !result.isConnected || !stats) return;
            if (stats.totalPurchases == null && stats.totalRefunds == null && !stats.email && stats.sellerNotes == null) return;
            const isCounting = stats.totalRefunds === -1;
            const updated = Object.assign({}, data, {
              totalPurchases: data.totalPurchases ?? stats.totalPurchases,
              totalRefunds:   !isCounting && stats.totalRefunds != null ? stats.totalRefunds : (data.totalRefunds ?? 0),
              // Item-quantity totals (used by the '전체 주문'/'전체 환불' ZD fields
              // instead of order counts) only finalize once refund-checking
              // completes (!isCounting) — same timing as totalRefunds, since both
              // are derived from the same per-order fetches.
              totalItemsPurchased: !isCounting && stats.totalItemsPurchased != null ? stats.totalItemsPurchased : data.totalItemsPurchased,
              totalItemsRefunded:  !isCounting && stats.totalItemsRefunded  != null ? stats.totalItemsRefunded  : data.totalItemsRefunded,
              // Known as early as the very first /orders-api/order/{id} fetch
              // (unlike the buyer-wide stats above), so merge it in regardless
              // of isCounting.
              sellerNotes: stats.sellerNotes ?? data.sellerNotes,
              buyer: Object.assign({}, data.buyer || {}, stats.email ? { BuyerEmail: stats.email } : {}),
            });
            // Always keep lastOrderData (what Auto-Fill reads) in sync with what's
            // rendered on screen — previously this assignment was gated behind
            // `!isCounting`, so clicking Auto-Fill while the panel still showed the
            // partial "counting" badge silently read a stale/undefined value instead
            // of the number the agent could actually see on screen (root cause of
            // the Auto-Fill popup showing a different 전체 주문 count than the panel).
            lastOrderData = Object.assign({}, lastOrderData, updated);
            if (!isCounting) {
              logStep_(`SC buyer stats: 구매 ${updated.totalItemsPurchased ?? stats.totalPurchases ?? data.totalPurchases ?? '?'}건 / 환불 ${stats.totalRefunds ?? '?'}건`);
            }
            // Prefer item quantity (matches the '✅전체 주문'/'❎전체 환불' ZD field
            // values) over raw order count — same precedence as renderOrder()'s
            // dispPurchases/dispRefunds. Previously always showed the raw order
            // count here, so a customer with 4 orders totaling 5 items showed 4.
            const tp = updated.totalItemsPurchased ?? updated.totalPurchases;
            const badge = result.querySelector('#sp-stat-badge');
            if (badge) {
              if (isCounting) {
                badge.innerHTML = tp != null ? `(구매 ${tp}건 / 환불 ${SPINNER_HTML})` : SPINNER_HTML;
              } else {
                const tr = updated.totalItemsRefunded ?? stats.totalRefunds ?? 0;
                badge.textContent = tp != null ? `(구매 ${tp}건 / 환불 ${tr}건)` : '';
              }
            }
            const linkWrap = result.querySelector('#sp-stat-link-wrap');
            if (linkWrap) {
              const newEmail = updated.buyer?.BuyerEmail;
              const newScUrl = sellerCentralSearchUrl_(updated.order?.SalesChannel, updated.address?.CountryCode, newEmail);
              if (isCounting && tp != null) {
                const valHtml = `구매 ${tp}건 / 환불 ${SPINNER_HTML}`;
                const cell = newScUrl
                  ? `<a href="${newScUrl}" target="_blank" rel="noopener" style="color:#5ba4cf;text-decoration:underline;font-weight:500;">${valHtml}</a>`
                  : `<span class="sp-val">${valHtml}</span>`;
                linkWrap.innerHTML = `<div class="sp-row"><span class="sp-label">구매이력 (2yr)</span>${cell}</div>`;
              } else {
                const tr = updated.totalItemsRefunded ?? stats.totalRefunds ?? 0;
                linkWrap.innerHTML = rowLinked('구매이력 (2yr)',
                  tp != null ? `구매 ${tp}건 / 환불 ${tr}건` : '—',
                  newScUrl);
              }
            }
            // Known as of the first response, unlike the buyer-stats fields
            // above — patch it in as soon as it arrives, not gated on isCounting.
            // Read-only mirror (see Seller Central's own order page to edit).
            // #sp-seller-notes-wrap starts out showing SPINNER_HTML (see
            // buildPanel()/resetPanel()) until this first real response lands.
            if (stats.sellerNotes != null) {
              const existingTa = document.getElementById('sp-seller-notes-content');
              if (existingTa) {
                existingTa.value = stats.sellerNotes;
              } else {
                const wrap = document.getElementById('sp-seller-notes-wrap');
                if (wrap) wrap.innerHTML = `<textarea id="sp-seller-notes-content" placeholder="(no seller notes)" readonly>${esc(stats.sellerNotes)}</textarea>`;
              }
            }
            maybeShowAutoFill(document.getElementById(PANEL_ID));
          });

          const prefs = getDataFetchPrefs();

          // Spigen-sheet product lookup is gated by the Product Info toggle.
          // Seller SKU / item ASINs are ORDER data and must NOT be gated by it.
          const doProductLookup = asins => {
            if (prefs.fetchProduct && asins && asins.length) {
              logStep_(`Product lookup: ${asins.join(', ')}`);
              renderAllProducts(asins);
            } else {
              _productReady = true;
              maybeShowAutoFill(document.getElementById(PANEL_ID));
            }
          };

          const rerenderOrder = renderAsin => {
            result.innerHTML = renderOrder(lastOrderData, orderId, renderAsin);
            result.querySelectorAll('.sp-block-title').forEach(t => {
              t.addEventListener('click', e => { e.stopPropagation(); t.closest('.sp-block').classList.toggle('collapsed'); });
            });
            applySectionState(result);
            maybeShowAutoFill(document.getElementById(PANEL_ID));
          };

          // Do we already have Seller SKU(s) from SP-API? (Almost never — items 403.)
          const haveSellerSku = (data.items || []).some(i => i.SellerSKU);

          if (haveSellerSku) {
            doProductLookup(itemAsins);
          } else {
            // No Seller SKU yet → ALWAYS query the Seller Central session for the full
            // item list (every SellerSKU + ASIN + price/tax + buyer proxy email),
            // regardless of itemsStatus or the Product Info toggle. Seller SKU is
            // order data the agent needs.
            logStep_(`Seller SKU: querying Seller Central (itemsStatus=${data.itemsStatus})…`);
            if (!resolvedAsin) {
              const asinValEl = result.querySelector('.sp-row .sp-val');
              if (asinValEl) asinValEl.textContent = 'Seller Central…';
            }
            fetchScItems(orderId, data.order?.SalesChannel, data.address?.CountryCode, (scItems, extras) => {
              if (_panelSession !== _session || !result.isConnected) return;
              if (scItems && scItems.length) {
                logStep_(`Seller SKU: SC returned ${scItems.length} item(s) — ${scItems.map(i => i.SellerSKU).filter(Boolean).join(', ') || 'no SKU'}`);
                lastOrderData.items = scItems;
                if (extras && extras.SenderEmail) {
                  lastOrderData.order = Object.assign({}, lastOrderData.order, { SenderEmail: extras.SenderEmail });
                }
                const newAsins = scItems.map(i => i.ASIN).filter(Boolean);
                if (asinInput && !asinInput.value) asinInput.value = newAsins.join(', ');
                rerenderOrder(newAsins.join(', ') || resolvedAsin);
                doProductLookup(newAsins.length ? newAsins : (resolvedAsin ? [resolvedAsin] : itemAsins));
              } else {
                // SC returned nothing → clear spinner → show '—'
                logStep_('Seller SKU: SC returned no items');
                const skuRow = result.querySelector('#sp-seller-sku-row');
                if (skuRow) skuRow.querySelector('.sp-val').innerHTML = '—';
                doProductLookup(resolvedAsin ? [resolvedAsin] : itemAsins);
              }
            });
          }
        } catch (err) {
          setStatus('Parse error: ' + err.message);
        }
      },
      onerror()   { setStatus('Cannot reach GAS endpoint — check GAS_URL in script settings.'); },
      ontimeout() {
        if (_panelSession !== _session) return;
        if (_retries < 2) {
          logStep_(`Order timeout — retry ${_retries + 1}/2…`);
          setStatus('Retrying…');
          setTimeout(() => fetchOrder(orderId, _retries + 1), 1000);
        } else {
          setStatus('Request timed out.');
          logStep_('Order fetch: timed out after 2 retries');
        }
      },
    }, { cacheTtl: ORDER_CLIENT_TTL, fresh: !!_fresh && !_retries });
  }

  function setStatus(msg, isLoading = false) {
    const html = isLoading ? `<span style="margin-right:6px;display:inline-block;">${SPINNER_HTML}</span>${esc(msg)}` : esc(msg);
    const el = document.getElementById('sp-status');
    if (el) { el.innerHTML = html; return; }
    const result = document.getElementById('sp-result');
    if (result) result.innerHTML = `<div id="sp-status">${html}</div>`;
  }

  function logStep_(msg) {
    const entries = document.getElementById('sp-log-entries');
    if (!entries) return;
    const t    = new Date().toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const line = document.createElement('div');
    line.textContent = `${t}  ${msg}`;
    entries.appendChild(line);
    entries.scrollTop = entries.scrollHeight;
  }

  // ── Auto-detect order IDs from visible ticket text ─────────────────────
  function detectOrderIds() {
    const m = location.pathname.match(/\/tickets\/(\d+)/);
    const pane = (m && document.querySelector(`[data-test-id="ticket-${m[1]}-standard-layout"]`)) || document.body;
    const inputText = [...pane.querySelectorAll('input, textarea')].map(el => el.value || '').join('\n');
    const text = (pane.innerText || '') + '\n' + inputText;
    return [...new Set([...text.matchAll(ORDER_RE)].map(m => m[1]))];
  }

  function updateDetectedChips(panel, skipAutoLoad, extraIds = []) {
    const domIds = detectOrderIds();
    // API-sourced IDs (from ticket description) take priority; DOM scan fills the rest
    const ids = [...new Set([...extraIds, ...domIds])];
    const bar = panel.querySelector('#sp-detected-ids');
    if (!bar) return;

    const existingSet = new Set([...bar.querySelectorAll('.sp-chip')].map(c => c.dataset.id));
    if (ids.length === existingSet.size && ids.every(id => existingSet.has(id))) return;
    // Don't wipe chips when an order is already loaded / being fetched
    if (ids.length === 0 && existingSet.size > 0) return;

    bar.innerHTML = '';
    if (ids.length >= 2) {
      const lbl = document.createElement('div');
      lbl.style.cssText = 'font-size:11px;color:#888;width:100%;margin-bottom:3px;font-weight:500;';
      lbl.textContent   = `주문 ID ${ids.length}개 발견 — 선택하세요:`;
      bar.appendChild(lbl);
    }
    ids.forEach(id => {
      const chip = document.createElement('span');
      chip.className    = 'sp-chip';
      chip.textContent  = id;
      chip.dataset.id   = id;
      chip.title        = 'Click to look up this order';
      chip.onclick = () => {
        document.getElementById('sp-order-input').value = id;
        fetchOrder(id);
      };
      bar.appendChild(chip);
    });

    if (!skipAutoLoad) {
      if (ids.length === 1 && document.getElementById('sp-status')) {
        const input = panel.querySelector('#sp-order-input');
        if (!input.value) { input.value = ids[0]; fetchOrder(ids[0]); }
      } else if (ids.length === 0) {
        setStatus('No Amazon order ID found on this ticket. Paste one above.');
      } else {
        setStatus('');
      }
    }
  }

  // ── Draggable panel ───────────────────────────────────────────────────────
  // Keep the floating panel's header on-screen so the ⚙ settings (and with it
  // "Dock in Apps panel") is always reachable. A saved top < 0 — from dragging
  // or a north-resize on a taller monitor — used to leave the header above the
  // viewport on a smaller/scaled screen (Windows 1080p report, 2026-10-07).
  function clampPanelIntoView_(panel) {
    if (!panel || !panel.isConnected || panel.classList.contains('sp-docked')) return;
    const r = panel.getBoundingClientRect();
    if (r.width === 0 && r.height === 0) return;
    const HEADER_MIN = 44, EDGE_MIN = 80;
    const maxTop  = Math.max(0, window.innerHeight - HEADER_MIN);
    const top     = Math.min(Math.max(0, r.top), maxTop);
    const minLeft = -(r.width - EDGE_MIN);
    const maxLeft = Math.max(0, window.innerWidth - EDGE_MIN);
    const left    = Math.min(Math.max(minLeft, r.left), maxLeft);
    if (Math.round(top) !== Math.round(r.top)) panel.style.top = Math.round(top) + 'px';
    if (Math.round(left) !== Math.round(r.left)) {
      panel.style.left = Math.round(left) + 'px';
      panel.style.right = 'auto';
    }
  }

  function makeDraggable(panel, handle) {
    handle.addEventListener('mousedown', e => {
      if (e.target.closest('#sp-abm-log-btn, #sp-settings-btn, #sp-minimize-btn, #sp-panel-close, .sp-re, button, input, textarea, select, a')) return;
      e.preventDefault();
      const rect = panel.getBoundingClientRect();
      const offX = e.clientX - rect.left;
      const offY = e.clientY - rect.top;
      let moved = false;
      const onMove = e2 => {
        if (!moved) panel.classList.add('sp-dragging');
        moved = true;
        panel.style.left  = (e2.clientX - offX) + 'px';
        panel.style.top   = (e2.clientY - offY) + 'px';
        panel.style.right = 'auto';
      };
      const onUp = () => {
        handle._dragMoved = moved;
        if (moved) {
          clampPanelIntoView_(panel);
          saveUi({ x: parseInt(panel.style.left), y: parseInt(panel.style.top) });
          panel.classList.remove('sp-dragging');
        }
        removeEventListener('mousemove', onMove);
        removeEventListener('mouseup', onUp);
      };
      addEventListener('mousemove', onMove);
      addEventListener('mouseup', onUp);
    });
  }

  // ── Resizable panel — 8 edge/corner grips bypass scrollbar interception ──
  // Root cause of old approach: panel.addEventListener('mousedown') was never
  // fired on the right edge because #sp-panel-body's native scrollbar
  // (overflow-y:auto) intercepts pointer events and does not bubble them.
  // Fix: invisible position:absolute divs (.sp-re-*) at z-index:9998 sit above
  // the scrollbar and fire reliable mousedown events directly.
  function makeResizable_(panel) {
    const MIN_W = 280, MAX_W = 700, MIN_H = 80;

    const DIRS = {
      'sp-re-n':  { top: true },
      'sp-re-ne': { top: true,    right: true  },
      'sp-re-e':  {               right: true  },
      'sp-re-se': { bottom: true, right: true  },
      'sp-re-s':  { bottom: true              },
      'sp-re-sw': { bottom: true, left: true   },
      'sp-re-w':  {               left: true   },
      'sp-re-nw': { top: true,    left: true   },
    };

    Object.entries(DIRS).forEach(([cls, h]) => {
      const el = panel.querySelector('.' + cls);
      if (!el) return;

      el.addEventListener('mousedown', e => {
        e.preventDefault();
        e.stopPropagation();

        const isDocked = panel.classList.contains('sp-docked');

        const r = panel.getBoundingClientRect();
        if (!isDocked) {
          // Normalize to left/top anchor so all resize math is consistent
          panel.style.left  = r.left + 'px';
          panel.style.right = 'auto';
          panel.style.top   = r.top  + 'px';
        }

        const startX = e.clientX, startY = e.clientY;
        const startLeft = r.left, startTop = r.top;
        const startW = r.width;
        // In docked mode use offsetHeight (unclipped CSS height) not BCR height
        // which is capped to the mount's visible area and causes wrong startH.
        const startH = isDocked ? panel.offsetHeight : r.height;

        const onMove = ev => {
          const dx = ev.clientX - startX;
          const dy = ev.clientY - startY;
          if (!isDocked) {
            if (h.right) panel.style.width  = Math.max(MIN_W, Math.min(MAX_W, startW + dx)) + 'px';
            if (h.left) {
              const nw = Math.max(MIN_W, Math.min(MAX_W, startW - dx));
              panel.style.width = nw + 'px';
              panel.style.left  = (startLeft + startW - nw) + 'px';
              panel.style.right = 'auto';
            }
            if (h.top) {
              const nh = Math.max(MIN_H, startH - dy);
              panel.style.height = nh + 'px';
              panel.style.top    = (startTop + startH - nh) + 'px';
            }
          }
          // Bottom-drag resizes height in both floating and docked modes
          if (h.bottom) {
            const newH = Math.max(MIN_H, startH + dy);
            panel.style.height = newH + 'px';
            // Docked: scroll the nearest scrollable ancestor so the resize grip
            // stays visible under the cursor regardless of which container clips.
            if (isDocked) el.scrollIntoView({ block: 'end', behavior: 'instant' });
          }
        };

        const onUp = () => {
          if (isDocked) {
            saveUi({ dockH: panel.offsetHeight });
          } else {
            clampPanelIntoView_(panel);
            saveUi({
              x: panel.style.left ? parseInt(panel.style.left) : null,
              y: panel.style.top  ? parseInt(panel.style.top)  : null,
              w: panel.offsetWidth,
              h: panel.offsetHeight,
            });
          }
          document.body.style.cursor     = '';
          document.body.style.userSelect = '';
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup',   onUp);
        };

        document.body.style.cursor     = getComputedStyle(el).cursor;
        document.body.style.userSelect = 'none';
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup',   onUp);
      });
    });
  }

  // ── Dock mode: mount the panel inside Zendesk's Apps panel ────────────────
  // Userscripts can't be native ZAF apps, so we DOM-inject. The mount point is
  // auto-discovered from the installed app iframes (ChannelReply/Trello live on
  // zdusercontent.com); we inject our panel as the first child of the scroll
  // container that holds them. A MutationObserver re-mounts after re-renders.
  let _dockObserver = null;
  let _dockObserverDebounce = null;
  let _dockFailedAt = null; // timestamp of first failed dock attempt after nav; used for floating fallback

  // Returns the DOM element for the current ticket panel, or null if not on a ticket.
  // Used to scope all DOM queries so we never mount into a stale/hidden ticket panel
  // (Zendesk keeps every open tab's DOM alive simultaneously).
  function ticketScope_() {
    const m = location.pathname.match(/\/tickets\/(\d+)/);
    const id = m && m[1];
    return (id && document.querySelector(`[data-test-id="ticket-${id}-standard-layout"]`)) || null;
  }

  function findAppsPanelMount_() {
    const W = window.innerWidth, H = window.innerHeight;

    // 0. Scope all queries to the current ticket's DOM so we never find/mount into
    //    a stale ticket panel — Zendesk keeps all open tab DOMs alive simultaneously.
    const _scope = ticketScope_() || document;

    // 1. Explicit Zendesk data-test-id selectors.
    // "omnipanel-pane-wrapper-apps" is confirmed by console diagnostics as the
    // correct Apps panel container in the new Zendesk navigation layout.
    const direct = _scope.querySelector(
      '[data-test-id="omnipanel-pane-wrapper-apps"], ' +
      '[data-test-id="ticket-apps-pane"], [data-test-id="apps-tray"], ' +
      '[data-test-id="omnipanel-apps"], [data-test-id="ticket_sidebar"], ' +
      '[data-test-id="apps-container"], [data-test-id="sidebar-apps"]'
    );
    if (direct) return direct;

    // 1.5. New Zendesk omnipanel — find the Apps content area.
    // The new Zendesk right panel has a narrow icon bar (~64px, contains
    // data-test-id="omnipanel-selector-item-*" buttons) and a wider content area
    // as its sibling. Walk up from the Apps icon button; at the first ancestor
    // where the icon bar (narrow ≤120px) and a wider content sibling both exist,
    // return the content sibling as the mount point.
    const omnipanelAppsBtn = _scope.querySelector('[data-test-id="omnipanel-selector-item-apps"]');
    if (omnipanelAppsBtn) {
      let el15 = omnipanelAppsBtn;
      for (let d = 0; d < 10 && el15.parentElement && el15 !== document.body; d++) {
        el15 = el15.parentElement;
        const visKids = [...el15.children].filter(c => {
          if (c.id === PANEL_ID) return false;
          const r = c.getBoundingClientRect();
          return r.width > 0 && r.height > 0;
        });
        if (visKids.length < 2) continue;
        const selBranch = visKids.find(c => c.contains(omnipanelAppsBtn));
        if (!selBranch) continue;
        const selW = selBranch.getBoundingClientRect().width;
        if (selW > 120) continue; // icon bar must be narrow
        const contentBranch = visKids.find(c => c !== selBranch && c.getBoundingClientRect().width > 80);
        if (contentBranch) {
          // Go deeper: the contentBranch may be a wrapper with multiple sections
          // in a flex-row layout. Find the apps-specific child where [data-app-id]
          // or zdusercontent iframes live so we mount at the same level as ChannelReply.
          const deepApp = contentBranch.querySelector('[data-app-id]');
          if (deepApp?.parentElement && deepApp.parentElement !== contentBranch) {
            logStep_('Dock: apps list found');
            return deepApp.parentElement;
          }
          const deepIframe = [...contentBranch.querySelectorAll('iframe')]
            .find(f => /zdusercontent\.com/.test(f.src || ''));
          if (deepIframe) {
            let iEl = deepIframe;
            for (let j = 0; j < 8 && iEl.parentElement && iEl !== contentBranch; j++) {
              iEl = iEl.parentElement;
              if (iEl.parentElement === contentBranch) break;
            }
            if (iEl !== contentBranch) {
              logStep_('Dock: apps iframe parent found');
              return iEl.parentElement || iEl;
            }
          }
          // Guard: only mount in the Apps pane. When Customer Context / Knowledge
          // / another section is active, contentBranch is their pane wrapper
          // (e.g. omnipanel-pane-wrapper-customer-context). Mounting there would
          // show GCX Reply in the wrong section and trigger churn.
          // Break out so mountDocked_ can auto-click the Apps icon instead.
          const _cbId = contentBranch.getAttribute('data-test-id') || '';
          if (_cbId && !_cbId.includes('apps')) break;

          logStep_('Dock: omnipanel content found');
          return contentBranch;
        }
      }
    }

    // 2. [data-app-id] filtered by CENTER position — Zendesk adds this attribute to
    //    each installed app wrapper (ChannelReply block, etc.). Using center-X instead
    //    of left edge is robust across different window widths. Center X > 60% of
    //    viewport means the element sits in the right portion of the page (Apps panel),
    //    not in the center ticket-editor area.
    const sidebarApp = [..._scope.querySelectorAll('[data-app-id]')].find(el => {
      const r = el.getBoundingClientRect();
      if (r.width === 0 || r.height === 0) return false;
      return (r.left + r.width / 2) > W * 0.6;
    });
    if (sidebarApp?.parentElement) {
      logStep_('Dock: data-app-id found');
      return sidebarApp.parentElement;
    }

    // Guard: if the omnipanel icon bar (the narrow button strip containing
    // omnipanel-selector-item-* buttons) is completely absent, Zendesk is mid-SPA
    // and has torn down the old ticket panel but not yet rendered the new one.
    // Any zdusercontent iframes found by Step 3 belong to the old ticket context
    // and will be replaced within milliseconds, making any mount instantly unstable.
    // Return null here so the observer/heartbeat keeps retrying until the omnipanel
    // is ready, rather than triggering the unstable-mount loop that resets
    // _dockFailedAt on every brief success and prevents the floating fallback.
    if (!_scope.querySelector('[data-test-id*="omnipanel-selector-item"]')) return null;

    // 3. zdusercontent iframe walk-up
    const appIframe = [..._scope.querySelectorAll('iframe')].find(f => /zdusercontent\.com/.test(f.src || ''));

    if (appIframe) {
      let el = appIframe;
      let appBlock = null;
      for (let i = 0; i < 22 && el.parentElement && el !== document.body; i++) {
        el = el.parentElement;
        const children = [...el.children].filter(c => c.id !== PANEL_ID);
        if (children.length < 2) continue;
        const hasIframeBranch = children.some(c => c === appIframe || c.contains(appIframe));
        const hasOtherBranch  = children.some(c => !c.contains(appIframe) && c.getBoundingClientRect().height > 0);
        if (hasIframeBranch && hasOtherBranch) { appBlock = el; break; }
      }
      if (appBlock?.parentElement) {
        logStep_('Dock: app-block found');
        return appBlock.parentElement;
      }
      el = appIframe;
      const scrollEls = [];
      for (let i = 0; i < 18 && el.parentElement && el !== document.body; i++) {
        el = el.parentElement;
        const oy = getComputedStyle(el).overflowY;
        if ((oy === 'auto' || oy === 'scroll') && el.clientHeight > 100) scrollEls.push(el);
      }
      if (scrollEls.length) {
        return scrollEls.reduce((b, c) => c.clientHeight > b.clientHeight ? c : b);
      }
      return appIframe.parentElement;
    }

    // 4. Semantic: <aside> or complementary role on the right side.
    //    Exclude modal dialogs (onboarding, announcements) and aria-hidden overlays —
    //    the new Zendesk navigation renders a data-test-id="onboarding-panel" aside
    //    on the right side that is hidden (aria-hidden=true, role=dialog).
    const aside = [...document.querySelectorAll('aside, [role="complementary"]')].find(el => {
      if (el.getAttribute('role') === 'dialog') return false;
      if (el.getAttribute('aria-hidden') === 'true') return false;
      if (el.hasAttribute('aria-modal')) return false;
      const r = el.getBoundingClientRect();
      return r.left > W * 0.4 && r.height > H * 0.3;
    });
    if (aside) return aside;

    // 5. "Apps" heading text — Zendesk always renders this in the apps panel header.
    const appsHeading = [...document.querySelectorAll('h1,h2,h3,h4,strong,b,span,label,div')]
      .find(el => {
        if (el.children.length > 0) return false;
        if (el.textContent.trim() !== 'Apps') return false;
        const r = el.getBoundingClientRect();
        return (r.left + r.width / 2) > W * 0.4 && r.top < H * 0.4;
      });
    if (appsHeading) {
      let el = appsHeading.parentElement;
      for (let i = 0; i < 15 && el && el !== document.body; i++) {
        const r = el.getBoundingClientRect();
        if (r.height > H * 0.4 && el.children.length >= 1) return el;
        el = el.parentElement;
      }
    }

    return null;
  }

  function mountDocked_(panel) {
    // Set up the DOM-watch observer BEFORE any early return so it is always
    // active — even when the Apps pane isn't rendered yet.
    // Debounced at 80ms to avoid calling mountDocked_() hundreds of times
    // per second when Zendesk is doing complex React re-renders.
    if (!_dockObserver) {
      _dockObserver = new MutationObserver(() => {
        if (!loadUi().dockMode || panel.isConnected) return;
        // Immediate (no debounce) check — Zendesk may render omnipanel-pane-wrapper-apps
        // only briefly during a React commit before moving it into the hub iframe.
        // Catching it here without delay is critical; the 80ms debounce below misses it.
        if ((ticketScope_() || document).querySelector('[data-test-id="omnipanel-pane-wrapper-apps"]')) {
          mountDocked_(panel);
          return;
        }
        clearTimeout(_dockObserverDebounce);
        _dockObserverDebounce = setTimeout(() => {
          if (!loadUi().dockMode || panel.isConnected) return;
          if ((ticketScope_() || document).querySelector('[data-test-id="omnipanel-pane-wrapper-apps"]')) {
            mountDocked_(panel);
          } else if (!panel._autoClickedAppsBtn) {
            // Apps pane not rendered — Apps section may be inactive.
            // Click the Apps icon to activate it. Allow re-click after 2s in
            // case the first click didn't trigger a render (React synthetic event issue).
            const appsBtn = (ticketScope_() || document).querySelector('[data-test-id="omnipanel-selector-item-apps"]');
            if (appsBtn && appsBtn.getAttribute('aria-pressed') !== 'true') {
              panel._autoClickedAppsBtn = true;
              logStep_('Dock: clicking Apps icon to activate section');
              appsBtn.click();
              setTimeout(() => { panel._autoClickedAppsBtn = false; }, 2000);
            }
          }
        }, 80);
      });
      _dockObserver.observe(document.body, { childList: true, subtree: true });
    }

    const mount = findAppsPanelMount_();
    if (!mount) {
      // If the Apps icon is in the DOM but its section isn't active (e.g. Customer
      // Context is the saved default), click it so Zendesk renders
      // omnipanel-pane-wrapper-apps. The immediate observer check will then catch
      // the DOM mutation and mount correctly.  2s guard prevents rapid re-clicks.
      if (!panel._autoClickedAppsBtn) {
        const appsBtn_ = (ticketScope_() || document).querySelector('[data-test-id="omnipanel-selector-item-apps"]');
        if (appsBtn_ && appsBtn_.getAttribute('aria-pressed') !== 'true') {
          panel._autoClickedAppsBtn = true;
          logStep_('Dock: clicking Apps icon to activate section');
          appsBtn_.click();
          setTimeout(() => { panel._autoClickedAppsBtn = false; }, 2000);
        }
      }
      if (!panel.isConnected) {
        if (loadUi().dockMode) {
          // After 2s of failed dock attempts, show panel floating so the user
          // always sees it even when Zendesk renders the omnipanel inside a
          // cross-origin hub iframe (unreachable from the main document).
          // The heartbeat's inFloatingFallback path keeps retrying mountDocked_()
          // so the panel re-docks the moment the omnipanel appears in main doc.
          if (!_dockFailedAt) _dockFailedAt = Date.now();
          if (Date.now() - _dockFailedAt >= 1000) {
            panel.classList.remove('sp-docked');
            panel._gcxHiddenBySection = false;
            panel.style.display = '';
            document.body.appendChild(panel);
            clampPanelIntoView_(panel);
            logStep_('Dock: 1s timeout — floating fallback (heartbeat retries dock)');
          } else {
            logStep_('Dock: mount not found — staying detached, observer + heartbeat will retry.');
          }
        } else {
          panel.classList.remove('sp-docked');
          document.body.appendChild(panel);
          logStep_('Dock: mount not found — showing floating, heartbeat will retry.');
        }
      }
      return;
    }
    _dockFailedAt = null; // reset — dock succeeded
    // Standalone block at the top of the Apps panel — independent of ChannelReply.
    if (panel.parentElement !== mount || panel.previousElementSibling) {
      mount.insertBefore(panel, mount.firstChild);
      const mountTag = mount.tagName + (mount.className ? '.' + mount.className.toString().trim().split(/\s+/)[0] : '');
      logStep_('Dock: mounted in ' + mountTag + ' (kids=' + mount.children.length + ')');
    }
    panel.classList.add('sp-docked');

    // Restore user-set docked height. Default is 300px so ChannelReply
    // remains visible below GCX Reply in the 735px wrapper.
    const savedDockH = loadUi().dockH;
    panel.style.height = (savedDockH > 80 ? savedDockH : 300) + 'px';

    // The new Zendesk omnipanel-pane-wrapper-apps container uses flex-direction:row,
    // which places GCX Reply and ChannelReply side-by-side instead of stacked.
    // Force column layout so apps stack vertically.
    // Scroll is handled by #sp-panel-body inside the panel (same as floating mode),
    // so we do not need a max-height or overflow on the mount itself.
    mount.style.flexDirection = 'column';
    mount.style.overflowX = 'hidden';
    panel.classList.remove('minimized');
    const tgl = document.getElementById('sp-toggle-btn');
    if (tgl) tgl.style.display = 'none';

    // New Zendesk omnipanel: the content area is shared across sections
    // (Apps, Customer Context, Knowledge, etc.). Sync GCX Reply's visibility
    // with the Apps icon button so it only shows when Apps section is active.
    // Always re-attach to the *current* Apps icon DOM element — Zendesk
    // replaces this element on every SPA navigation, making the old observer stale.
    const appsIconBtn_ = (ticketScope_() || document).querySelector('[data-test-id="omnipanel-selector-item-apps"]');
    if (appsIconBtn_) {
      const syncSectionVis_ = () => {
        const live = appsIconBtn_.isConnected;
        const inApps = appsIconBtn_.getAttribute('aria-pressed') === 'true';
        logStep_('Dock: syncSection inApps=' + inApps + ' btn=' + (live ? 'live' : 'STALE'));
        panel._gcxHiddenBySection = !inApps;
        panel.style.display = inApps ? '' : 'none';
      };
      syncSectionVis_();
      // When the Apps section isn't currently active, the panel is hidden.
      // Auto-click the Apps icon once per ticket visit so it appears without
      // the user having to manually switch sections. Step 3 of findAppsPanelMount_
      // finds the zdusercontent iframe mount even when Apps isn't active, so the
      // old auto-click code in the !mount branch is never reached in this case.
      if (panel._gcxHiddenBySection && !panel._autoClickedAppsBtn) {
        panel._autoClickedAppsBtn = true;
        logStep_('Dock: clicking Apps icon (section not active)');
        appsIconBtn_.click();
        setTimeout(() => { panel._autoClickedAppsBtn = false; }, 2000);
      }
      // Always recreate the observer so the callback closure captures the current
      // appsIconBtn_ reference. Reusing the old observer across SPA navigations
      // would keep the stale closure from the first mountDocked_() call, which
      // reads a detached element's aria-pressed and permanently hides the panel.
      if (panel._sectionObserver) panel._sectionObserver.disconnect();
      panel._sectionObserver = new MutationObserver(syncSectionVis_);
      panel._sectionObserver.observe(appsIconBtn_, { attributes: true, attributeFilter: ['aria-pressed'] });
    } else {
      panel._gcxHiddenBySection = false;
      panel.style.display = '';
      if (panel._sectionObserver) { panel._sectionObserver.disconnect(); panel._sectionObserver = null; }
    }
  }

  function mountFloating_(panel) {
    if (_dockObserver) { _dockObserver.disconnect(); _dockObserver = null; }
    if (panel._sectionObserver) { panel._sectionObserver.disconnect(); panel._sectionObserver = null; }
    panel._gcxHiddenBySection = false;
    panel.style.display = '';
    panel.classList.remove('sp-docked');
    if (panel.parentElement !== document.body) document.body.appendChild(panel);
    // Restore saved floating geometry.
    const ui = loadUi();
    panel.style.width  = (ui.w || 330) + 'px';
    if (ui.h) panel.style.height = ui.h + 'px';
    if (ui.x != null) { panel.style.left = ui.x + 'px'; panel.style.right = 'auto'; }
    else { panel.style.left = 'auto'; panel.style.right = '16px'; }
    panel.style.top = (ui.y != null ? ui.y : 72) + 'px';
    clampPanelIntoView_(panel);
    logStep_('Dock: restored floating panel');
  }

  // ── Settings drawer (data fetch preferences) ─────────────────────────────
  function initSettings_(panel) {
    const btn    = panel.querySelector('#sp-settings-btn');
    const drawer = panel.querySelector('#sp-settings-drawer');
    if (!btn || !drawer) return;

    const uiState = loadUi();
    const prefs = getDataFetchPrefs();

    drawer.innerHTML = `
      <div class="sp-settings-inner">
        <div style="border-bottom:1px solid rgba(0,0,0,0.06);padding-bottom:8px;margin-bottom:8px;">
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;font-size:12px;color:#1c1c1e;margin-bottom:0;">
            <input type="checkbox" id="sp-dock-chk" ${uiState.dockMode === true ? 'checked' : ''}/>
            Dock in Apps panel
          </label>
        </div>
        <div>
          <div style="font-size:11px;color:#888;margin-bottom:6px;font-weight:500;">Data Fetch</div>
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;font-size:12px;color:#1c1c1e;margin-bottom:5px;">
            <input type="checkbox" id="sp-fetch-order-chk" ${prefs.fetchOrder ? 'checked' : ''}/>
            Order Info
          </label>
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;font-size:12px;color:#1c1c1e;margin-bottom:5px;">
            <input type="checkbox" id="sp-fetch-shipping-chk" ${prefs.fetchShipping ? 'checked' : ''}/>
            Shipping Address
          </label>
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;font-size:12px;color:#1c1c1e;margin-bottom:5px;">
            <input type="checkbox" id="sp-fetch-product-chk" ${prefs.fetchProduct ? 'checked' : ''}/>
            Product Info
          </label>
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;font-size:12px;color:#1c1c1e;margin-bottom:5px;">
            <input type="checkbox" id="sp-notes-toggle"/>
            Notes
          </label>
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;font-size:12px;color:#1c1c1e;margin-bottom:5px;">
            <input type="checkbox" id="sp-seller-notes-toggle"/>
            Seller Notes
          </label>
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;font-size:12px;color:#1c1c1e;">
            <input type="checkbox" id="sp-show-log-chk" ${!uiState.logCollapsed ? 'checked' : ''}/>
            Debug log
          </label>
        </div>
        <div style="border-top:1px solid rgba(0,0,0,0.06);padding-top:8px;margin-top:8px;">
          <div style="font-size:11px;color:#888;margin-bottom:6px;font-weight:500;">Alerts</div>
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;font-size:12px;color:#1c1c1e;">
            <input type="checkbox" id="sp-abm-alerts-chk" ${abmAlertsEnabled_() ? 'checked' : ''}/>
            ABM 전송 알림 (팝업/배지)
          </label>
          <div style="font-size:10.5px;color:#999;margin-top:3px;line-height:1.4;">
            끄면 자동 팝업 없이 헤더의 ABM 버튼 숫자로만 미전송 건을 표시합니다.
          </div>
        </div>
        <div style="border-top:1px solid rgba(0,0,0,0.06);padding-top:8px;margin-top:8px;">
          <div style="font-size:11px;color:#888;margin-bottom:6px;font-weight:500;">MCF 담당자</div>
          <select id="sp-mcf-person-select" style="width:100%;padding:4px 6px;font-size:12px;border:1px solid #d8dcde;border-radius:6px;box-sizing:border-box;">
            <option value="">선택 안 함</option>
            ${MCF_PERSON_OPTIONS.map(p => `<option value="${p}" ${getMcfPerson_() === p ? 'selected' : ''}>${p}</option>`).join('')}
          </select>
          <div style="font-size:10.5px;color:#999;margin-top:3px;line-height:1.4;">
            MCF 발송 로그에 담당자로 기록됩니다.
          </div>
        </div>
      </div>
    `;

    const dockChk = drawer.querySelector('#sp-dock-chk');
    if (dockChk) dockChk.addEventListener('change', () => {
      saveUi({ dockMode: dockChk.checked });
      if (dockChk.checked) mountDocked_(panel);
      else mountFloating_(panel);
    });

    const fetchOrderChk = drawer.querySelector('#sp-fetch-order-chk');
    const fetchShippingChk = drawer.querySelector('#sp-fetch-shipping-chk');
    const fetchProductChk = drawer.querySelector('#sp-fetch-product-chk');

    function onDataFetchToggleChange() {
      const prefs = {
        fetchOrder: fetchOrderChk.checked,
        fetchShipping: fetchShippingChk.checked,
        fetchProduct: fetchProductChk.checked,
      };
      saveDataFetchPrefs(prefs);
      const result = panel.querySelector('#sp-result');
      if (result && lastOrderData) {
        const orderId = panel.querySelector('#sp-order-input')?.value.trim() || '';
        const panelAsin = panel.querySelector('#sp-asin-input')?.value.trim() || '';
        result.innerHTML = renderOrder(lastOrderData, orderId, panelAsin);
        result.querySelectorAll('.sp-block-title').forEach(title => {
          title.addEventListener('click', e => {
            e.stopPropagation();
            title.closest('.sp-block').classList.toggle('collapsed');
          });
        });
        applySectionState(result);
      }
      const productResult = panel.querySelector('#sp-product-result');
      if (productResult) {
        productResult.style.display = prefs.fetchProduct ? 'block' : 'none';
      }
    }

    fetchOrderChk.addEventListener('change', onDataFetchToggleChange);
    fetchShippingChk.addEventListener('change', onDataFetchToggleChange);
    fetchProductChk.addEventListener('change', onDataFetchToggleChange);

    btn.addEventListener('click', e => {
      e.stopPropagation();
      drawer.classList.toggle('sp-settings-open');
    });

    // Apply saved log-collapsed state on open
    const logWrap = panel.querySelector('#sp-load-log');
    if (logWrap && uiState.logCollapsed) logWrap.classList.add('sp-log-collapsed');

    const showLogChk = drawer.querySelector('#sp-show-log-chk');
    if (showLogChk && logWrap) {
      showLogChk.addEventListener('change', () => {
        logWrap.classList.toggle('sp-log-collapsed', !showLogChk.checked);
        saveUi({ logCollapsed: !showLogChk.checked });
      });
    }

    const abmAlertsChk = drawer.querySelector('#sp-abm-alerts-chk');
    if (abmAlertsChk) {
      abmAlertsChk.addEventListener('change', () => {
        saveUi({ abmAlertsOff: !abmAlertsChk.checked });
        // Re-derive the badge immediately: turning alerts off removes any
        // currently-floating badge; turning them back on re-shows it if
        // undelivered replies exist.
        updateAbmBadge_();
      });
    }

    const mcfPersonSelect = drawer.querySelector('#sp-mcf-person-select');
    if (mcfPersonSelect) {
      mcfPersonSelect.addEventListener('change', () => {
        setMcfPerson_(mcfPersonSelect.value);
      });
    }

    const abmLogBtn = panel.querySelector('#sp-abm-log-btn');
    if (abmLogBtn) {
      abmLogBtn.addEventListener('click', e => {
        e.stopPropagation();
        showAbmPanel_();
      });
      // Populate the pending count as soon as the panel mounts, instead of
      // waiting up to 5 min for the next background sweep tick.
      updateAbmBadge_();
    }
  }

  // ── Init ──────────────────────────────────────────────────────────────────
  function isTicketPage_() { return !!location.pathname.match(/\/tickets\/\d+/); }
  function isFiltersPage_() { return !!location.pathname.match(/\/agent\/filters/); }

  function init() {
    if (_panelEl) return; // heartbeat may call init() before setTimeout(init,800) fires — block double-init
    if (!isTicketPage_() && !isFiltersPage_()) return;
    if (document.getElementById(PANEL_ID)) return;

    let toggleBtn = document.getElementById('sp-toggle-btn');
    if (!toggleBtn) {
      toggleBtn = document.createElement('button');
      toggleBtn.id          = 'sp-toggle-btn';
      toggleBtn.textContent = 'Order Lookup';
      document.body.appendChild(toggleBtn);
    }

    const panel = buildPanel();
    _panelEl = panel;
    const _savedUi = loadUi();
    const _startDocked = _savedUi.dockMode === true;

    if (_startDocked) {
      // Dock mode: never show the floating popup first. Mark docked immediately and
      // mount into the Apps panel below; until then the node stays detached (hidden),
      // so it never flashes as a floating window.
      panel.classList.add('sp-docked');
    } else {
      document.body.appendChild(panel);
      // Restore saved size + position (floating only).
      // Soft-clamp so ≥50px of the header stays grabbable; any other off-screen ok.
      if (_savedUi.w) panel.style.width  = _savedUi.w + 'px';
      if (_savedUi.h) panel.style.height = _savedUi.h + 'px';
      if (_savedUi.x != null) {
        panel.style.left  = _savedUi.x + 'px';
        panel.style.right = 'auto';
      }
      if (_savedUi.y != null) panel.style.top = _savedUi.y + 'px';
      clampPanelIntoView_(panel);
    }

    // Start minimized on filter/list pages; expanded on ticket pages
    if (!isTicketPage_()) panel.classList.add('minimized');

    const header = panel.querySelector('#sp-panel-header');
    makeDraggable(panel, header);
    if (!window._gcxClampOnResize) {
      window._gcxClampOnResize = true;
      addEventListener('resize', () => { if (_panelEl) clampPanelIntoView_(_panelEl); });
    }
    // Body is NOT draggable so its text stays selectable/copyable; drag via the title bar.
    makeResizable_(panel);

    initSettings_(panel);

    // Dock from the start (no floating flash). mountDocked_ retries internally
    // (and the heartbeat re-mounts) until the Apps panel is available.
    if (_startDocked) mountDocked_(panel);

    // Minimize / expand
    panel.querySelector('#sp-minimize-btn').onclick = e => {
      e.stopPropagation();
      const minimized = panel.classList.toggle('minimized');
      if (minimized) {
        panel.dataset.savedH = panel.style.height || '';
        panel.style.height = '';
      } else if (panel.dataset.savedH) {
        panel.style.height = panel.dataset.savedH;
      }
    };

    // Docked collapse/expand button (Garden-style chevron)
    const dockCollapseBtn = panel.querySelector('#sp-dock-collapse-btn');
    if (dockCollapseBtn) {
      dockCollapseBtn.onclick = e => {
        e.stopPropagation();
        const collapsed = panel.classList.toggle('minimized');
        dockCollapseBtn.setAttribute('aria-label', collapsed ? 'Expand section' : 'Collapse section');
      };
    }

    header.addEventListener('click', e => {
      if (header._dragMoved) { header._dragMoved = false; return; }
      if (e.target.closest('#sp-abm-log-btn, #sp-settings-btn, #sp-minimize-btn, #sp-panel-close, #sp-dock-collapse-btn, button')) return;
      if (panel.classList.contains('sp-docked')) {
        const collapsed = panel.classList.toggle('minimized');
        if (dockCollapseBtn) dockCollapseBtn.setAttribute('aria-label', collapsed ? 'Expand section' : 'Collapse section');
      } else if (panel.classList.contains('minimized')) {
        panel.classList.remove('minimized');
        if (panel.dataset.savedH) panel.style.height = panel.dataset.savedH;
      }
    });

    // Reactive compact layout via ResizeObserver
    if (typeof ResizeObserver !== 'undefined') {
      new ResizeObserver(([e]) => {
        panel.classList.toggle('sp-compact', e.contentRect.width < 260);
      }).observe(panel);
    }

    // Initialize product result visibility based on preferences
    const prefs = getDataFetchPrefs();
    const productResult = panel.querySelector('#sp-product-result');
    if (productResult) {
      productResult.style.display = prefs.fetchProduct ? 'block' : 'none';
    }

    panel.querySelector('#sp-panel-close').onclick = () => {
      panel.remove();
      toggleBtn.style.display = 'block';
    };
    toggleBtn.onclick = () => {
      toggleBtn.style.display = 'none';
      init();
    };

    const orderInput = panel.querySelector('#sp-order-input');
    const asinInput  = panel.querySelector('#sp-asin-input');

    panel.querySelector('#sp-lookup-btn').onclick = () => {
      const id = orderInput.value.trim();
      if (id) fetchOrder(id, 0, true); // manual Lookup = always fresh data
    };
    orderInput.addEventListener('keydown', e => { if (e.key === 'Enter') panel.querySelector('#sp-lookup-btn').click(); });

    panel.querySelector('#sp-product-btn').onclick = () => {
      const raw = asinInput.value.trim().toUpperCase();
      if (!raw) return;
      const asins = raw.split(',').map(a => a.trim()).filter(Boolean);
      renderAllProducts(asins, false, false, true); // manual click = always fresh from GAS
    };
    asinInput.addEventListener('keydown', e => { if (e.key === 'Enter') panel.querySelector('#sp-product-btn').click(); });

    panel.querySelector('#sp-autofill-btn').onclick = () => autoFillTicket(panel);
    panel.querySelector('#sp-mcf-btn').onclick = () => sendToMCF(panel);
    panel.querySelector('#sp-nrn-btn').onclick = () => markNoResponseNeeded(panel);

    // Mirror ChannelReply's button state: enable our NRN button only while CR's
    // Per-row copy buttons (delegated, guarded once). Copies the row's value text.
    if (!window.__gcxCopyInit) {
      window.__gcxCopyInit = true;
      const copyText_ = text => {
        if (navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(text);
        const ta = document.createElement('textarea');
        ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta); ta.focus(); ta.select();
        try { document.execCommand('copy'); } finally { ta.remove(); }
        return Promise.resolve();
      };
      document.addEventListener('click', e => {
        const btn = e.target.closest && e.target.closest('.sp-copy');
        if (!btn) return;
        e.preventDefault(); e.stopPropagation();
        // data-copy attribute wins (used by item-name buttons that aren't in .sp-row)
        let text = btn.dataset.copy ? btn.dataset.copy.trim() : '';
        if (!text) {
          const rowEl = btn.closest('.sp-row');
          const valEl = rowEl && (rowEl.querySelector('.sp-val') || rowEl.querySelector('a'));
          text = (valEl ? valEl.textContent : '').trim();
        }
        if (!text || text === '—') return;
        copyText_(text).then(() => {
          btn.classList.add('copied');
          btn.innerHTML = '✓';
          setTimeout(() => { btn.classList.remove('copied'); btn.innerHTML = COPY_SVG; }, 1000);
        }).catch(() => {});
      }, true);
    }

    // NRN button state — direct Seller Central check (see refreshNrnState_
    // above). Primary trigger is the order-load hook; this interval is just a
    // safety net that resets/re-checks the button on ticket (URL) change.
    if (!window.__gcxNrnInit) {
      window.__gcxNrnInit = true;
      let _lastUrl = location.href;
      refreshNrnState_();
      autoFillAbmOrderId_();
      autoCorrectDetectedLanguage_();
      setInterval(() => {
        if (location.href === _lastUrl) return;
        _lastUrl = location.href;
        const btn = document.querySelector('#sp-order-panel #sp-nrn-btn');
        if (btn) { btn.disabled = true; btn.title = 'Checking…'; }
        refreshNrnState_();
        autoFillAbmOrderId_();
        autoCorrectDetectedLanguage_();
      }, 1000);
    }

    const notesToggle  = panel.querySelector('#sp-notes-toggle');
    const notesSection = panel.querySelector('#sp-notes-section');
    const notesContent = panel.querySelector('#sp-notes-content');

    // A ticket page can render MULTIPLE customer-context cards (e.g. related
    // contacts, not just the requester), each with its own Notes textarea
    // sharing the exact same data-test-id — document.querySelector grabbed
    // whichever came first in the DOM, which is often a different, unrelated,
    // empty person's notes rather than the actual requester's. Disambiguate by
    // matching the card containing the requester's email — first against the
    // buyer email GCX Reply already resolved via SP-API (no extra fetch),
    // falling back to the ticket's own "via" from-address (fetchTicketJson_ is
    // already cached elsewhere) — and cache which element that is so the live
    // input listener below only reacts to edits on that specific textarea.
    let _notesTa = null;
    // The notes textarea lives inside Zendesk's "Customer context" omnipanel
    // tab, which — same as the requester-name card used for the Auto-Fill
    // Customer Full Name fallback — is lazy-rendered and doesn't exist in the
    // DOM at all until that tab has been opened at least once. Without this,
    // querySelectorAll below finds zero matches whenever the agent hasn't
    // already clicked into Customer context themselves, and both reading and
    // writing notes silently no-op.
    function resolveNotesTextarea_(cb) {
      ensureCustomerContextPanelOpen_().then(async switchedTab => {
        // The tab being "open" (aria-pressed=true) doesn't mean its content
        // has actually rendered yet — ensureCustomerContextPanelOpen_'s fixed
        // 700ms wait was frequently too short for the card/textarea to exist
        // in the DOM at all yet (confirmed live: querySelectorAll returned []
        // immediately after opening, but the same query found the real
        // element once given more time). Poll for the element to actually
        // appear instead of querying once.
        const all = (await pollUntil_(() => {
          const found = [...document.querySelectorAll('[data-test-id="notes-edit-text-area-test-id"]')];
          return found.length ? found : null;
        })) || [];
        proceed_(all);

        function proceed_(all) {
          const restoreTab_ = () => {
            if (!switchedTab) return;
            const appsBtn = document.querySelector('[data-test-id="omnipanel-selector-item-apps"]');
            if (appsBtn && appsBtn.getAttribute('aria-pressed') !== 'true') appsBtn.click();
          };

          if (all.length <= 1) { _notesTa = all[0] || null; restoreTab_(); cb(_notesTa); return; }

          const matchByEmail = email => {
            if (!email) return null;
            return all.find(ta => {
              const card = ta.closest('[data-test-id="customer-context-card-body-test-id"]');
              return card && (card.textContent || '').includes(email);
            }) || null;
          };

          const byOrder = matchByEmail(lastOrderData?.buyer?.BuyerEmail);
          if (byOrder) { _notesTa = byOrder; restoreTab_(); cb(_notesTa); return; }

          const ticketId = location.pathname.match(/\/tickets\/(\d+)/)?.[1];
          if (!ticketId) { _notesTa = all[0]; restoreTab_(); cb(_notesTa); return; }
          fetchTicketJson_(ticketId, t => {
            const byVia = matchByEmail(t?.via?.source?.from?.address);
            _notesTa = byVia || all[0]; // last resort: old (first-match) behavior
            restoreTab_();
            cb(_notesTa);
          });
        }
      });
    }

    function refreshNotes() {
      resolveNotesTextarea_(ta => {
        if (!ta) {
          // Distinguish "couldn't check at all" from "checked, genuinely
          // empty" — the omnipanel sidebar itself occasionally fails to
          // render (a Zendesk-side long-session SPA-state issue, confirmed
          // fixed by a full page reload; not something GCX Reply can force).
          // Silently showing the same "(no notes)" placeholder as a real
          // empty result would mislead the agent into thinking a customer
          // has no notes when GCX Reply simply couldn't look.
          notesContent.value = '';
          notesContent.placeholder = '⚠ 불러오기 실패 — 페이지 새로고침 후 재시도';
          return;
        }
        notesContent.placeholder = '(no notes)'; // restore normal placeholder now that resolution succeeded
        // The Customer context card/textarea can already exist in the DOM
        // (satisfying resolveNotesTextarea_ above) before Zendesk's own async
        // fetch has actually populated its value — reading immediately after
        // opening the tab was catching it empty even when real notes exist.
        // Poll briefly instead of reading once.
        pollUntil_(() => ta.value || null).then(val => { notesContent.value = val || ''; });
      });
    }
    notesToggle.addEventListener('change', () => {
      notesSection.style.display = notesToggle.checked ? 'block' : 'none';
      if (notesToggle.checked) refreshNotes();
      saveUi({ notes: notesToggle.checked });
    });
    if (loadUi().notes) {
      notesToggle.checked = true;
      notesSection.style.display = 'block';
      refreshNotes();
    }

    // Seller Notes toggle — same simple show/hide pattern as Notes above.
    // Unlike Notes, the underlying data comes from an already-in-flight SC
    // fetch (fetchScBuyerStats_, shared with the 구매이력 stats) rather than a
    // live DOM read, so toggling this on just reveals whatever's already been
    // patched into #sp-seller-notes-content (or the disabled/empty state if
    // that fetch hasn't resolved yet).
    const sellerNotesToggle  = panel.querySelector('#sp-seller-notes-toggle');
    const sellerNotesSection = panel.querySelector('#sp-seller-notes-section');
    sellerNotesToggle.addEventListener('change', () => {
      sellerNotesSection.style.display = sellerNotesToggle.checked ? 'block' : 'none';
      saveUi({ sellerNotes: sellerNotesToggle.checked });
    });
    if (loadUi().sellerNotes) {
      sellerNotesToggle.checked = true;
      sellerNotesSection.style.display = 'block';
    }

    // Read-only mirrors only — Notes edits belong on the real Zendesk
    // Customer context textarea, Seller Notes edits belong on the real
    // Seller Central order page. GCX Reply just displays both.
    // Zendesk re-renders a NEW textarea DOM node after a save (same
    // data-test-id, different element instance) — a plain reference-equality
    // guard (e.target !== _notesTa) would silently reject every edit after
    // the first, since _notesTa keeps pointing at the old, now-detached
    // node. Self-heal by dropping stale references so the next input event
    // re-adopts whichever (still-connected) element fired it.
    document.addEventListener('input', e => {
      if (!notesToggle.checked) return;
      if (!e.target.matches('[data-test-id="notes-edit-text-area-test-id"]')) return;
      if (_notesTa && !_notesTa.isConnected) _notesTa = null;
      if (_notesTa && e.target !== _notesTa) return; // edit on a different person's notes — ignore
      if (!_notesTa) _notesTa = e.target; // no resolution done yet, or previous ref went stale
      notesContent.value = e.target.value;
    });

    // ── Persist section collapse state (capture phase runs before stopPropagation) ──
    panel.addEventListener('click', e => {
      const title = e.target.closest('.sp-block-title');
      if (!title) return;
      const block = title.closest('[data-sp-section]');
      if (!block) return;
      setTimeout(() => {
        const c = loadUi().collapsed || {};
        c[block.dataset.spSection] = block.classList.contains('collapsed');
        saveUi({ collapsed: c });
      }, 0);
    }, true);

    // ── Reset panel on ticket navigation ────────────────────────────────────
    function resetPanel() {
      orderInput.value = '';
      asinInput.value  = '';
      lastOrderData    = null;
      lastProductData  = null;
      lastAmazonProduct = null;
      _productReady    = false;
      _ticketJsonCache = null;
      _cachedBodyText  = null; // force a fresh reflow read for the new ticket
      document.getElementById('sp-sc-session-warn')?.remove();
      // Only wipe GCX Reply-filled ZD fields when Auto-Fill was actually confirmed on this
      // ticket. Clearing unconditionally caused Zendesk's own saved field values to be lost
      // whenever the agent navigated away and back (regression from v2.7.3).
      _gcrFilledThisTicket = false;
      const result = document.getElementById('sp-result');
      if (result) result.innerHTML = '<div id="sp-status">Scanning ticket for order IDs…</div>';
      const productResult = document.getElementById('sp-product-result');
      if (productResult) productResult.innerHTML = '';
      showAiReasonBtn_('');
      lastAiReason = null;
      const chips = document.getElementById('sp-detected-ids');
      if (chips) chips.innerHTML = '';
      const autoBar = panel.querySelector('#sp-autofill-bar');
      if (autoBar) autoBar.style.display = 'none';
      const mcfBar = panel.querySelector('#sp-mcf-bar');
      if (mcfBar) mcfBar.style.display = 'none';
      const notesToggleEl = panel.querySelector('#sp-notes-toggle');
      if (notesToggleEl) notesToggleEl.checked = false;
      const notesSectionEl = panel.querySelector('#sp-notes-section');
      if (notesSectionEl) notesSectionEl.style.display = 'none';
      // Back to the loading spinner for the new ticket's fetch — otherwise the
      // previous ticket's Seller Notes text would stay visible (stale) until
      // the new fetchScBuyerStats_ call resolves.
      const sellerNotesWrap = document.getElementById('sp-seller-notes-wrap');
      if (sellerNotesWrap) sellerNotesWrap.innerHTML = `<div class="sp-row"><span class="sp-val">${SPINNER_HTML}</span></div>`;
      const logEntries = document.getElementById('sp-log-entries');
      if (logEntries) logEntries.innerHTML = '';
      setFillStatus(panel, '');
      _panelSession++;
    }

    // ── Collapse raw ABM Amazon-template comments when a clean copy exists ──
    // ABM_TicketMerge (server-side, separate GAS project) posts a clean copy
    // of the buyer's real ABM message right after any raw "You have received
    // a message" Amazon-template comment, but deliberately leaves the raw
    // original fully intact — redacting it broke Zendesk's own native
    // ticket-merge quoting (a merged ticket's system note reads a comment's
    // CURRENT stored text; a redacted comment made that note unreadable, see
    // ABM_TicketMerge v19). Purely visual/client-side here: collapses the raw
    // comment to a one-line toggle so agents see only the clean message +
    // attachment by default, without touching any Zendesk data.
    //
    // Never mutates the article's own children — Zendesk's React tree owns
    // those (e.g. a "Show more" truncation toggle can re-render them at any
    // time, which would silently wipe anything we inserted inside). Only
    // toggles a CSS class on the article itself (safe className mutation) and
    // adds a small toggle button as a SIBLING, never a child.
    const ABM_RAW_MARKERS_ = [
      'you have received a message',
      'resolve case',
      'report questionable activity',
    ];
    // Primary signal: Amazon's raw ABM notification template wraps the
    // buyer's own message in an identically-styled <pre> block across every
    // marketplace/language — the exact same locale-agnostic anchor
    // ABM_TicketMerge's server-side cleanAbmMessageText_ already relies on
    // (see ABM_RAW_PRE_RE in ABM_TicketMerge/Code.js). The English text
    // markers below only ever match English-language tickets — confirmed
    // live that Japanese (#1000155119) and French (#1000155061) raw
    // templates contain ZERO of these English phrases (fully translated:
    // "メッセージが届きました" / "Vous avez reçu un message", etc.), so
    // isRawAbmComment_ silently never recognized them as ABM comments at
    // all and they never collapsed. The clean copy comment ABM_TicketMerge
    // posts alongside is always plain text with no <pre> — confirmed live,
    // zero false-positive risk of the pair matching each other.
    function isRawAbmComment_(el) {
      if (el.querySelector('pre[style*="-o-pre-wrap"]')) return true;
      const text = (el.innerText || '').toLowerCase();
      return ABM_RAW_MARKERS_.filter(m => text.includes(m)).length >= 2;
    }
    // A Zendesk end-user identity merge (ABM_TicketMerge's
    // normalizeAbmRequesterIdentity_, server-side) can leave the raw
    // comment's cached author name shorter/staler than the clean copy's —
    // e.g. "Thomas" (the merged-away, now-inactive user's name at the time
    // it was cached) vs "Thomas Doubleday" (the canonical profile's fuller
    // name, enriched after the merge). Confirmed live on ticket #1000155360:
    // exact-string pairing failed even though a real clean copy existed
    // right next to the raw comment. Treat a whole-word prefix match either
    // direction as the same person, not just a byte-identical string.
    function abmAuthorNamesMatch_(a, b) {
      if (!a || !b) return false;
      const an = a.trim().toLowerCase();
      const bn = b.trim().toLowerCase();
      if (an === bn) return true;
      return an.startsWith(bn + ' ') || bn.startsWith(an + ' ');
    }
    function collapseRawAbmComments_() {
      // Zendesk's aria-label used to START with "Message from {name}, ..." —
      // it now PREPENDS the full message preview text first, with "Message
      // from {name}, via {channel}, {time}" appended at the END instead
      // (confirmed live on ticket #1000157352, e.g. "メッセージが届きました。
      // ...(long body)... Message from 鈴木, via email, Yesterday 5:31 pm").
      // The `^=` "starts with" selector below (and both `^Message from`
      // regexes) silently matched ZERO articles ever since Zendesk made that
      // change — nothing threw, the feature just stopped collapsing anything
      // on any ticket, with no error to notice. Switched to `*=`/unanchored
      // matching, which works regardless of where "Message from" lands.
      const articles = Array.from(document.querySelectorAll('article[aria-label*="Message from"]'));
      articles.forEach((article, idx) => {
        if (article.dataset.gcxAbmScanned) return;
        if (!isRawAbmComment_(article)) { article.dataset.gcxAbmScanned = '1'; return; }

        const label = article.getAttribute('aria-label') || '';
        const author = (label.match(/Message from ([^,]+),/) || [])[1];
        if (!author) return; // can't confirm a pair — leave fully visible

        // Only collapse if THIS raw comment's OWN clean copy exists — scoped
        // to same-author non-raw articles strictly BETWEEN this one and the
        // next raw article (from anyone), not "anywhere in the ticket."
        // ABM_TicketMerge posts one clean copy per raw message, so on an
        // ongoing back-and-forth each raw/clean pair is its own independent
        // unit — cleanup can legitimately succeed for some of a buyer's
        // messages and fail for others (a transient failure, not a bug in
        // this collapse logic). The old "any match anywhere" check meant
        // once ONE of a buyer's messages had a real clean pair, EVERY later
        // raw message from that same buyer got collapsed too, even ones
        // whose own cleanup never ran — confirmed live on ticket
        // #1000161639 ("Theo", 4 raw messages: #1/#2/#4 correctly had their
        // own clean copies, #3 ("I have a photo... but I don't have the
        // id") had none, immediately followed by an agent reply — yet it
        // still collapsed by matching against #2's or #4's unrelated pair).
        // Intervening AGENT replies between a raw message and its clean
        // copy are fine (cleanup can lag a reply) — only hitting the NEXT
        // RAW message means this one's own pair was never found.
        let hasCleanPair = false;
        for (let i = idx + 1; i < articles.length; i++) {
          const other = articles[i];
          if (isRawAbmComment_(other)) break; // next raw message starts — no pair was found before it
          const otherAuthor = ((other.getAttribute('aria-label') || '').match(/Message from ([^,]+),/) || [])[1];
          if (abmAuthorNamesMatch_(author, otherAuthor)) { hasCleanPair = true; break; }
        }
        if (!hasCleanPair) return;

        article.dataset.gcxAbmScanned = '1';
        article.classList.add('gcx-abm-raw-collapsed');

        const toggle = document.createElement('div');
        toggle.className = 'gcx-abm-raw-toggle';
        toggle.textContent = '▶ Show original Amazon email';
        toggle.addEventListener('click', (e) => {
          e.stopPropagation();
          const collapsed = article.classList.toggle('gcx-abm-raw-collapsed');
          toggle.textContent = collapsed ? '▶ Show original Amazon email' : '▼ Hide original Amazon email';
        });
        if (article.parentElement) article.parentElement.insertBefore(toggle, article);
      });
    }

    function autoDetectAll() {
      const _session = _panelSession;
      autoCorrectDetectedLanguage_(); // order-less trigger: Country* may already be set (e.g. ABM auto-fill)
      getTicketFields((orderId, asin, bodyIds) => {
        if (_panelSession !== _session) return; // stale callback from prev ticket — discard
        const orderInput = panel.querySelector('#sp-order-input');
        if (orderId && orderInput && !orderInput.value) {
          // Custom field has order ID → use it directly, chips are informational only
          orderInput.value = orderId;
          fetchOrder(orderId);
          updateDetectedChips(panel, true);
        } else {
          // No custom field order ID → merge message-body IDs with DOM scan
          // If exactly 1 total → auto-fetch; if multiple → show chips for user to pick
          updateDetectedChips(panel, false, bodyIds);
        }

        const detectedAsin = asin || [...new Set([...getCachedBodyText_().matchAll(ASIN_RE)].map(m => m[1]))][0];
        if (detectedAsin) {
          const ai = document.getElementById('sp-asin-input');
          if (ai && !ai.value) {
            ai.value = detectedAsin;
            const _asinOnly = !orderId && (!bodyIds || !bodyIds.length);
            const _prefs = getDataFetchPrefs();
            if (_prefs.fetchProduct) {
              renderAllProducts([detectedAsin], false, _asinOnly);
            }
          }
        }
      });
    }

    let lastTicketId = location.pathname.match(/\/tickets\/(\d+)/)?.[1];
    let navTimer = null;
    let _mountNavTimer = null;
    function onNav() {
      const newId = location.pathname.match(/\/tickets\/(\d+)/)?.[1];
      if (newId) {
        // Navigated to a ticket — always expand
        if (panel.classList.contains('minimized')) {
          panel.classList.remove('minimized');
          if (panel.dataset.savedH) panel.style.height = panel.dataset.savedH;
        }
        if (newId !== lastTicketId) {
          lastTicketId = newId;
          resetPanel();
          prefetchTicketData_(); // v3.8.0: network starts now; render flow below unchanged
          _dockFailedAt = null; // restart the 2s floating-fallback timer for the new ticket
          clearTimeout(navTimer);
          navTimer = setTimeout(() => { autoDetectAll(); collapseRawAbmComments_(); }, 1500);
          // Re-run mountDocked_() after each ticket navigation so that the section
          // observer re-attaches to the fresh DOM elements Zendesk creates on SPA
          // render. Without this, the stale observer never fires when the user
          // clicks the Apps icon after a ticket switch, leaving the panel hidden.
          if (loadUi().dockMode) {
            // Reset auto-click flag so the observer can click the Apps icon
            // once on this new ticket if the Apps section isn't the default.
            panel._autoClickedAppsBtn = false;
            clearTimeout(_mountNavTimer);
            _mountNavTimer = setTimeout(() => mountDocked_(panel), 500);
          }
        }
      } else {
        // Left ticket pages (filters, views, etc.) — always collapse
        lastTicketId = null;
        if (location.pathname.startsWith('/agent/filters')) resetPanel();
        if (!panel.classList.contains('minimized')) {
          panel.dataset.savedH = panel.style.height || '';
          panel.style.height = '';
          panel.classList.add('minimized');
        }
      }
    }
    const origPush    = history.pushState.bind(history);
    const origReplace = history.replaceState.bind(history);
    history.pushState    = (...a) => { origPush(...a);    onNav(); };
    history.replaceState = (...a) => { origReplace(...a); onNav(); };
    window.addEventListener('popstate', onNav);

    let scanTimer = null;
    const observer = new MutationObserver(() => {
      clearTimeout(scanTimer);
      scanTimer = setTimeout(() => {
        updateDetectedChips(panel, !!panel.querySelector('#sp-order-input')?.value);
        collapseRawAbmComments_();
      }, 1200);
    });
    observer.observe(document.body, { childList: true, subtree: true });

    if (isTicketPage_()) prefetchTicketData_();
    if (isTicketPage_()) setTimeout(() => { autoDetectAll(); collapseRawAbmComments_(); }, 1500);

    showUpdatePopupIfNeeded_();
  }

  // ── Embedded MCF page autofill (runs on Seller Central MCF pages) ──────────
  // This makes GCX Reply self-sufficient for MCF autofill — "Amazon MCF Autofill"
  // Tampermonkey script is no longer required (kept only as emergency clipboard fallback).
  function initMcfPage_() {
    const sl = ms => new Promise(r => setTimeout(r, ms));

    // --- Kat form helpers ---
    function setKat(el, val) {
      if (!el || val == null) return false;
      try {
        el.value = val;
        el.setAttribute('value', val);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        const inner = el.shadowRoot?.querySelector('input,textarea');
        if (inner) {
          inner.value = val;
          inner.dispatchEvent(new Event('input', { bubbles: true }));
          inner.dispatchEvent(new Event('change', { bubbles: true }));
        }
        return true;
      } catch(e) { return false; }
    }

    function setByAnyLbl(labels, v) {
      if (!v) return false;
      const el = [...document.querySelectorAll('kat-input')].find(k =>
        labels.some(l => (k.getAttribute('label') || '').trim().toLowerCase() === l.toLowerCase())
      );
      return el ? setKat(el, v) : false;
    }

    function setCountry(code) {
      if (!code) return false;
      const upper = code.toUpperCase().replace(/^UK$/, 'GB').replace(/^EL$/, 'GR');
      const dd = document.querySelector('kat-dropdown[label="Country"]') ||
                 document.querySelector('kat-dropdown[label="국가"]');
      if (!dd?.shadowRoot) return false;
      const sr = dd.shadowRoot;
      const header = sr.querySelector('.select-header, [part="dropdown-header"]');
      if (header) header.click();
      let attempts = 0;
      const timer = setInterval(() => {
        const opt = sr.querySelector(`kat-option[value="${upper}"]`);
        if (opt && opt.offsetParent !== null) {
          const inner = opt.shadowRoot?.querySelector('.content-wrapper');
          (inner || opt).click();
          clearInterval(timer);
          return;
        }
        if (++attempts > 30) clearInterval(timer);
      }, 100);
      return true;
    }

    function fillAll({ name, street, city, state, postal, phone, email, asin }) {
      let ok = false;
      ok = setByAnyLbl(['Full name', '전체 이름'], name)           || ok;
      ok = setByAnyLbl(['Street address', '상세 주소'], street)    || ok;
      ok = setByAnyLbl(['City', '도시'], city)                     || ok;
      ok = setByAnyLbl(['State / Province', '시/도'], state)       || ok;
      ok = setByAnyLbl(['Postcode', '우편번호'], postal)           || ok;
      ok = setByAnyLbl(['Phone number', '전화번호'], phone)        || ok;
      if (email) ok = setByAnyLbl(['Email address', '이메일 주소'], email) || ok;
      if (asin) {
        const inp = document.getElementById('sku-search-input');
        if (inp) {
          setKat(inp, asin);
          ok = true;
          setTimeout(() => {
            const inner = inp.shadowRoot?.querySelector('input');
            if (inner) {
              inner.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true, composed: true }));
              inner.dispatchEvent(new KeyboardEvent('keyup',   { key: 'Enter', keyCode: 13, bubbles: true, composed: true }));
            }
          }, 400);
        }
      }
      return ok;
    }

    // --- SKU auto-select (highest fulfillable count, skip amzn.* internal SKUs) ---
    function autoSelectBestSku() {
      let attempts = 0;
      const timer = setInterval(() => {
        const components = [...document.querySelectorAll('.search-result-component')];
        if (!components.length) { if (++attempts > 60) clearInterval(timer); return; }
        // Quantity badge always leads with the number (e.g. "1,234 fulfillable" in
        // English, "192 주문 처리 가능" in Korean) — match the leading digits instead
        // of the localized unit text so this works regardless of SC UI language.
        const entries = components.map(comp => {
          const m = (comp.querySelector('.search-result-component-quantity')?.textContent || '').match(/^([\d,]+)/);
          return { count: m ? parseInt(m[1].replace(/,/g, ''), 10) : 0, comp };
        }).filter(({ comp }) => !/\bamzn[.\-]/i.test(comp.textContent || ''));
        if (!entries.length) { if (++attempts > 60) clearInterval(timer); return; }
        entries.sort((a, b) => b.count - a.count);
        const xBtn = entries[0].comp.querySelector('.search-result-x');
        if (!xBtn) { if (++attempts > 60) clearInterval(timer); return; }
        clearInterval(timer);
        // React fiber click (most reliable)
        const fKey = Object.keys(xBtn).find(k => k.startsWith('__reactFiber') || k.startsWith('__reactInternalInstance'));
        if (fKey) {
          let fiber = xBtn[fKey];
          while (fiber) {
            if (fiber.memoizedProps && typeof fiber.memoizedProps.onClick === 'function') {
              fiber.memoizedProps.onClick({ preventDefault(){}, stopPropagation(){}, type:'click', target:xBtn });
              return;
            }
            fiber = fiber.return;
          }
        }
        ['pointerover','pointerenter','mouseover','mouseenter','pointermove','mousemove',
         'pointerdown','mousedown','pointerup','mouseup','click'].forEach(type =>
          xBtn.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, composed: true, view: window }))
        );
      }, 500);
    }

    // --- Standard shipping ---
    function forceStandard() {
      let left = 80;
      const timer = setInterval(() => {
        const grp = document.querySelector('kat-radiobutton-group[name="shipping-speed"]') ||
                    document.querySelector('kat-radiobutton-group');
        if (grp) {
          const rb = [...document.querySelectorAll('kat-radiobutton')].find(r => (r.getAttribute('value') || '').toLowerCase() === 'standard');
          if (rb) { rb.click(); }
          else { try { grp.value = 'Standard'; grp.dispatchEvent(new Event('change', { bubbles: true })); } catch(_) {} }
          if ((grp.value || '').toLowerCase() === 'standard') { clearInterval(timer); return; }
        }
        const labels = [...document.querySelectorAll('kat-label[part="radiobutton-label"], kat-label[for]')];
        const std = labels.find(el =>
          /\bstandard\b/i.test(el.getAttribute('text') || '') ||
          /\bstandard\b/i.test(el.textContent || '')
        );
        if (std) { const nl = std.querySelector('label[for]'); if (nl) { nl.click(); clearInterval(timer); return; } }
        if (--left <= 0) clearInterval(timer);
      }, 350);
    }

    // --- Order ID ---
    const _isOrderLbl = lbl => lbl.includes('order id') || lbl.includes('merchant order id') || lbl.includes('주문 id');
    function isOrderIdFilled_() {
      const k = [...document.querySelectorAll('kat-input')].find(k => _isOrderLbl((k.getAttribute('label') || '').trim().toLowerCase()));
      if (k && (k.value || k.getAttribute('value'))) return true;
      const inner = document.querySelector('input[name*="orderId"], input[id*="orderId"]');
      return !!(inner && inner.value.trim());
    }
    function setOrderId_(v) {
      if (!v) return;
      const k = [...document.querySelectorAll('kat-input')].find(k => _isOrderLbl((k.getAttribute('label') || '').trim().toLowerCase()));
      if (k) { setKat(k, v); return; }
      const inner = document.querySelector('input[name*="orderId"], input[id*="orderId"]');
      if (inner) { inner.value = v; inner.dispatchEvent(new Event('input', { bubbles: true })); }
    }

    const _MCF_GAS_EP = 'https://script.google.com/macros/s/AKfycbwM02GYF6gvdT1mSD7ePeLMU2huRz4ARl2E5AJ2Oh-nKYLWD3nbyHqAcNreM8wGZwdo/exec';
    // Was native fetch() with no retry — Apps Script's exec URL 302-redirects
    // internally, and CORS on that redirected response is flaky from a
    // third-party origin (sellercentral.amazon.*), so this failed silently and
    // permanently on any hiccup. GM_xmlhttpRequest bypasses CORS entirely
    // (same as every other GAS call in this script), plus a few retries for
    // plain network/timeout/cold-start failures.
    // `person` is resolved right at the call site (GM_getValue, prompting via
    // showMcfPersonPicker_ if unset) — falls back to the backend's own
    // '김지우' default (MCF_Tracking/tamperMonkey.js markMcfRow_) only if the
    // agent dismissed the prompt without picking.
    async function markRowMcf_(email, person) {
      if (!email) return null;
      const url = _MCF_GAS_EP + '?email=' + encodeURIComponent(email) + '&action=markMcf&match=last&person=' + encodeURIComponent(person || '');
      for (let attempt = 0; attempt < 3; attempt++) {
        const res = await gmGet_(url);
        if (res && res.status === 200) {
          try {
            const data = JSON.parse(res.responseText);
            return (data?.success === true && data.orderId) ? data.orderId : null;
          } catch (_) { /* malformed response — fall through to retry */ }
        }
        if (attempt < 2) await sl(1000);
      }
      return null;
    }

    // --- Price warning: red-highlight + note bubble if MCF fee > 1/3 item price ---
    function parsePriceNum_(str) {
      if (!str) return null;
      const m = (str || '').match(/([\d]+[.,][\d]{2}|[\d]+)/);
      return m ? parseFloat(m[1].replace(',', '.')) : null;
    }

    function applyPriceWarning_(el, deliveryPrice, itemPrice) {
      if (!el || deliveryPrice == null || !itemPrice || itemPrice <= 0) return;
      const ratio = deliveryPrice / itemPrice;
      if (ratio <= 1 / 3) {
        el.style.cssText = '';
        el._spBubble?.remove();
        delete el._spBubble;
        return;
      }
      el.style.cssText = 'background:#c0392b!important;color:#fff!important;padding:1px 5px;border-radius:3px;font-weight:700;';
      if (!el._spBubble) {
        const bubble = document.createElement('span');
        bubble.style.cssText = 'display:inline-flex;align-items:center;justify-content:center;margin-left:5px;cursor:default;background:#c0392b;color:#fff;border-radius:50%;width:15px;height:15px;font-size:10px;font-weight:700;vertical-align:middle;position:relative;';
        bubble.textContent = '!';
        const tip = document.createElement('div');
        tip.style.cssText = 'display:none;position:absolute;bottom:120%;left:50%;transform:translateX(-50%);background:#c0392b;color:#fff;padding:6px 10px;border-radius:6px;font-size:11px;white-space:nowrap;z-index:99999;box-shadow:0 2px 8px rgba(0,0,0,.35);pointer-events:none;';
        tip.textContent = `MCF 배송비가 상품가의 1/3 초과 (${Math.round(ratio * 100)}%)`;
        bubble.appendChild(tip);
        bubble.addEventListener('mouseenter', () => { tip.style.display = 'block'; });
        bubble.addEventListener('mouseleave', () => { tip.style.display = 'none'; });
        el.parentElement?.insertBefore(bubble, el.nextSibling);
        el._spBubble = bubble;
      } else {
        const tip = el._spBubble.querySelector('div');
        if (tip) tip.textContent = `MCF 배송비가 상품가의 1/3 초과 (${Math.round(ratio * 100)}%)`;
      }
    }

    function watchPriceWarning_(itemPrice) {
      // Disconnect before DOM changes to prevent characterData mutations from tip.textContent
      // re-triggering the observer in an infinite loop that locks the MCF page UI.
      const observer = new MutationObserver(() => {
        observer.disconnect();
        document.querySelectorAll('.total-price-label').forEach(el =>
          applyPriceWarning_(el, parsePriceNum_(el.textContent), itemPrice)
        );
        observer.observe(document.body, { childList: true, subtree: true, characterData: true });
      });
      document.querySelectorAll('.total-price-label').forEach(el =>
        applyPriceWarning_(el, parsePriceNum_(el.textContent), itemPrice)
      );
      observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    }

    // --- Main hash/window.name bridge autofill ---
    async function autoFillFromHash_() {
      // window.name is the primary bridge (persists through Netlify → SC cross-origin redirects)
      let wnEncoded = '';
      const wn = window.name || '';
      if (wn.startsWith('spigen_mcf:')) {
        wnEncoded = wn.slice('spigen_mcf:'.length);
        window.name = '';
      }
      // Fallback: sessionStorage hash saved by MCF Autofill script (if also installed)
      const storedHash = sessionStorage.getItem('_spigen_mcf_hash') || '';
      const encoded = wnEncoded || (storedHash.includes('spigen_mcf=') ? storedHash.split('spigen_mcf=')[1] : '');
      if (!encoded) return;

      try {
        const d = JSON.parse(decodeURIComponent(atob(encoded)));
        if (!d || d.region === 'JP') return;
        sessionStorage.removeItem('_spigen_mcf_hash');

        // Fill form — retry until address fields appear (some SC markets show them only after SKU selection)
        let filled = false;
        let asinKicked = false;
        for (let i = 0; i < 60; i++) {
          await sl(500);
          const ok = fillAll({ name: d.name, street: d.street, city: d.city, state: d.state,
                               postal: d.postal, phone: d.phone, email: d.email, asin: d.asin });
          if (ok) { filled = true; break; }
          if (!asinKicked && d.asin) {
            const searchInner = document.getElementById('sku-search-input')?.shadowRoot?.querySelector('input');
            if (searchInner?.value?.trim()) { asinKicked = true; autoSelectBestSku(); }
          }
        }
        if (!filled) return;

        if (d.asin && !asinKicked) autoSelectBestSku();
        if (d.country) setTimeout(() => setCountry(d.country), 800);

        if (d.email) {
          // Ask right here, on the SC MCF page itself — not on the Zendesk
          // side — so it fires no matter how the agent got to this page (the
          // panel's "→ MCF" button, a Netlify-redirected marketplace link, OR
          // a Zendesk Macro's hyperlink straight to this URL — all three
          // funnel through this same autoFillFromHash_ call once the hash
          // payload is present). GM_getValue/GM_setValue, not saveUi, so the
          // same saved pick is visible across every sellercentral.amazon.*
          // domain and the Zendesk ⚙ settings drawer alike.
          let mcfPerson = getMcfPerson_();
          if (!mcfPerson) {
            const picked = await showMcfPersonPicker_();
            if (picked) { mcfPerson = picked; setMcfPerson_(picked); }
          }
          const orderId = await markRowMcf_(d.email, mcfPerson);
          // Single-shot write here used to race the Order ID kat-input's own
          // render timing — autoSelectBestSku()/setCountry() above are still
          // running async (SKU search can take many seconds), and if that
          // input isn't in the DOM yet at this exact moment, setOrderId_
          // silently finds nothing and never retries. The sheet side already
          // has the correct 담당자/발송일자/Status/Q-col order ID by this
          // point (confirmed live in "MCF 발송 로그") — only the on-page
          // field was ever losing it. Retry on the same cadence as the
          // "wait for Standard shipping" loop below instead of one attempt.
          if (orderId) {
            let setAttempts = 150;
            const orderTimer = setInterval(() => {
              if (isOrderIdFilled_()) { clearInterval(orderTimer); return; }
              setOrderId_(orderId);
              if (--setAttempts <= 0) clearInterval(orderTimer);
            }, 400);
          }
        }

        // Wait for item + order ID to be ready, then select Standard
        let left = 150;
        const shTimer = setInterval(() => {
          if (isOrderIdFilled_()) { forceStandard(); clearInterval(shTimer); return; }
          if (--left <= 0) clearInterval(shTimer);
        }, 400);

        // Price warning
        if (d.itemPrice) watchPriceWarning_(parseFloat(d.itemPrice));
      } catch(e) { /* silent */ }
    }

    // Run after DOM is ready
    if (document.body) {
      setTimeout(autoFillFromHash_, 500);
    } else {
      const t = setInterval(() => { if (document.body) { clearInterval(t); setTimeout(autoFillFromHash_, 500); } }, 20);
    }

    // SPA hashchange (if Amazon routes to MCF via client-side nav without full reload)
    window.addEventListener('hashchange', () => {
      if (location.hash?.includes('spigen_mcf=')) {
        try { sessionStorage.setItem('_spigen_mcf_hash', location.hash); } catch(e) {}
        history.replaceState(null, '', location.pathname + location.search);
        autoFillFromHash_();
      }
    });
  }

  // 국가 코드 → MCF 랜딩 URL
  // JP: 셀러 센트럴 재팬 직접 이동
  // GB: UK 마켓플레이스 전환용 netlify (chipper)
  // 그 외 EU / 기타: 비UK EU 전환용 netlify (dulcet)
  // netlify 페이지가 #spigen_mcf= 해시를 읽어 적절한 SC 마켓플레이스로 리다이렉트
  function getMcfBase_(country) {
    if (country === 'JP') return 'https://sellercentral-japan.amazon.com/mcf/orders/create-order';
    if (country === 'GB') return 'https://chipper-youtiao-04b62d.netlify.app';
    return 'https://dulcet-cendol-0ec85d.netlify.app';
  }

  // MCF 링크 패치 — event delegation instead of per-link onclick.
  // Per-link onclick requires the link to already be in the DOM when patchMcfLinks_
  // runs. Zendesk iframes are detected by a 1-second interval, so clicks in the first
  // second fall through to the raw Netlify redirect (no hash). Event delegation on
  // the document/iframe catches ALL clicks regardless of when the link was added,
  // including future links injected by Zendesk's SPA renderer.
  function patchMcfLinks_(rootEl) {
    const root = rootEl instanceof Document ? rootEl : (rootEl ? rootEl.ownerDocument || document : document);
    if (root._mcfDelegated) return;
    root._mcfDelegated = true;

    root.addEventListener('click', async (e) => {
      const link = e.target.closest('a');
      if (!link) return;
      const href = link.href || '';
      const isNetlify = href.includes('.netlify.app');
      const isMcfDirect = href.includes('mcf/orders/create-order') && !href.includes('spigen_mcf=');
      if (!isNetlify && !isMcfDirect) return;

      e.preventDefault();
      e.stopPropagation();

      try {
        const ctxEmail = await getCustomerContextEmail_();
        const payload = buildMcfPayload_(document.getElementById(PANEL_ID), ctxEmail);
        const encoded = btoa(encodeURIComponent(JSON.stringify(payload)));
        const base = isMcfDirect
          ? href.split('#')[0].replace(/\?[^]*$/, '')
          : getMcfBase_(payload.country);
        const _w = window.open(base + '#spigen_mcf=' + encoded, '_blank');
        if (_w) { try { _w.name = 'spigen_mcf:' + encoded; } catch(e) {} }
      } catch (err) {
        // Encoding failed — open the original link so the click isn't silently swallowed.
        window.open(href, '_blank');
      }
    }, true); // capture phase: fires before React/Zendesk handlers
  }

  // 메인 문서 감시 (delegation attaches once; MutationObserver no longer needed for links)
  patchMcfLinks_(document);

  // iframe 내부도 감시 (Zendesk 에디터 등)
  // patchMcfLinks_ now uses event delegation so we just need to attach the listener
  // to the iframe document once — no MutationObserver per-link needed anymore.
  if (typeof setInterval === 'function') setInterval(() => {
    document.querySelectorAll('iframe').forEach(iframe => {
      try {
        const doc = iframe.contentDocument;
        if (!doc || !doc.body) return;
        patchMcfLinks_(doc);
      } catch(e) {}
    });
  }, 1000);

  // Heartbeat: if both the panel AND the toggle button disappear from the DOM
  // (Zendesk SPA re-render, slow initial load, etc.) re-run init automatically.
  // typeof guards make this file safe to copy verbatim into GAS (where setInterval/setTimeout are undefined).
  if (typeof setInterval === 'function') setInterval(() => {
    // Dock mode: Zendesk's React re-render detaches our injected node on ticket
    // navigation. Keep the SAME panel element and re-insert whenever needed.
    // Also retry from floating-fallback (panel on body while dockMode=true) in
    // case the Apps panel flyout was closed when we first tried to mount.
    if (loadUi().dockMode === true && _panelEl) {
      const inFloatingFallback = _panelEl.isConnected &&
        _panelEl.parentElement === document.body &&
        !_panelEl.classList.contains('sp-docked');
      // Don't remount when GCX Reply is intentionally hidden because the user
      // has switched to a different omnipanel section (Customer Context, etc.).
      const hiddenBySection = _panelEl._gcxHiddenBySection === true;
      // Detect wrong-ticket mount: panel is connected but lives inside a
      // different ticket's DOM (Zendesk keeps all open tabs in DOM simultaneously).
      const _hbScope = ticketScope_();
      const inWrongTicket = _panelEl.isConnected && !!_hbScope && !_hbScope.contains(_panelEl);
      // Always re-mount when disconnected (detached from DOM), regardless of
      // hiddenBySection — the panel must be in the mount to be shown when the
      // user switches back to the Apps section.
      if (!_panelEl.isConnected || (!hiddenBySection && _panelEl.offsetHeight === 0) || inFloatingFallback || inWrongTicket) {
        if (inWrongTicket) {
          _panelEl.remove();
          logStep_('Dock: wrong ticket DOM, detaching and remounting...');
        }
        mountDocked_(_panelEl);
      }
      return;
    }
    if (!document.getElementById(PANEL_ID) && !document.getElementById('sp-toggle-btn')) {
      init();
    }
  }, 600);

  if (typeof setTimeout === 'function') setTimeout(init, 800);
})();


# Tampermonkey Scripts — Spigen GCX

Tampermonkey userscripts for the Spigen GCX Amazon operations workflow.

---

## Scripts

### GCX Reply (`v3.5.2`)
**Matches:** `spigenhelp.zendesk.com/agent/tickets/*`

Floating panel on Zendesk tickets. Fetches Amazon order data and Spigen product info, then fills all relevant ticket fields in one click. Also handles Amazon Buyer Message (ABM) replies and "No Response Needed" directly against Seller Central — no ChannelReply dependency.

**Order lookup:**
- Auto-detects order ID from Zendesk custom fields, ticket description, and page text
- Shows clickable chips when multiple order IDs are found; auto-fetches if only one is found
- Displays: Order ID (linked to Seller Central), Order Status, Purchase Date, Amount, Delivery Level, Ship Date, Fulfillment Channel, Ship Service Level, Buyer Name
- Shipping Address (collapsible)
- Items list with SKU, quantity, and title
- Return ASIN (linked to Amazon product page)
- 구매이력 (2yr): total ITEM quantity purchased/refunded across all of the buyer's orders in the last 2 years (not raw order count — a customer with 4 orders totaling 5 items shows 5), linked to Seller Central order search

**Product info:**
- Auto-detects ASIN from Zendesk custom fields or page text
- Looks up: SKU, 모델명, 브랜드, 제조사명, 기종명, 색상명, 대분류, 생산업체, 원산지정보
- Data source priority: ASIN Master (Sheet1) → market sheet (Sheet2) → Amazon product page (fallback scrape)
- 판매 마켓 badges showing which marketplaces carry the product
- ASIN Sources panel (collapsed by default): shows Sheet1, Sheet2, and Amazon data side by side

**Auto-Fill Form button** (appears after order data loads):
- Opens an editable "Confirm Auto-Fill" popup — every row matches its real Zendesk field type (searchable dropdown for tagger/dropdown fields, editable text input for free-text fields), so an agent can correct a wrong or missing value before submitting, not just accept/reject it
- Fills all Zendesk custom fields via the Zendesk API in one click on confirm:
  Order ID, ASIN, 문의SKU, Customer Name, Purchase Date, Order Status, Order Total, Delivery Level, Country, Point of Purchase, Brand(상세), Device, Product Name, Fulfillment, ❗사진/영상 유무❗, ✅전체 주문, ❎전체 환불
- Device and Product Name fields are matched to dropdown options using token similarity, with a candidate picker shown when genuinely ambiguous
- A "what's new" popup shows once per version bump after Auto-Fill/panel changes ship

**→ MCF button** (appears after order data loads):
- Opens the Amazon MCF create-order page (global or JP) in a new tab
- Pre-fills recipient name, address, ASIN, and order ID via URL hash, picked up by the MCF Autofill script
- Also runs its own embedded MCF-page autofill (SKU search + auto-select highest-fulfillable-quantity result, shipping speed, Order ID lookup) independent of the standalone MCF Autofill script below

**Amazon Buyer Message (ABM) handling:**
- "Mark as NRN" button calls Seller Central's own internal messaging API directly (no ChannelReply) to mark a case resolved — works on any ABM ticket, order or not (case resolved from the buyer proxy address's embedded case ID)
- Public replies to ABM tickets are auto-relayed to the matching Seller Central case via the agent's own SC session; a reconciliation sweep catches replies that failed to relay (no SC session, page not open, etc.)
- Relay success/failure toasts and the floating "N replies not yet delivered" badge can be muted per-agent (⚙ → Alerts → ABM 전송 알림); the panel header's static **ABM** button always shows the undelivered count and opens the full relay-status log on demand — per-ticket links, per-row Retry / Mark delivered, checkbox selection with "Retry selected" (plus select-all) alongside "Retry all"
- Persistent banner warns when a required marketplace's Seller Central login session is missing, with direct login links

**Panel UX:**
- Draggable (grab header) and resizable (drag bottom-right corner)
- Minimize / close buttons; "Order Lookup" toggle button to reopen
- Compact layout mode auto-activates when panel width < 260px
- Auto-resets when navigating between tickets (SPA-aware)
- Load log at the bottom shows live fetch steps for debugging
- Settings drawer (⚙): Liquid Glass effect toggle, per-section data-fetch toggles (Order/Shipping/Product Info)

**Fallbacks:**
- If SP-API lacks buyer PII permission, fetches buyer email and 2-year order/item stats from the agent's existing Seller Central session
- If SP-API GetOrderItems is blocked (403), queries Seller Central orders-api for items using the agent's SC session

---

### Amazon MCF Autofill (`v1.4.3`)
**Matches:** `sellercentral.amazon.*` and `sellercentral-europe.amazon.*` — MCF create-order pages

Injects a floating panel on the MCF order creation page (EU marketplaces: UK, DE, FR, IT, ES). Autofills recipient name, address, and line items from order data passed by GCX Reply's MCF button (via URL hash, with a clipboard-paste fallback), reducing manual data entry when placing Multi-Channel Fulfillment orders.

- Auto-selects the SKU search result with the highest fulfillable quantity, skipping internal `amzn.*` SKUs — quantity is parsed from the leading digit count in the result badge so this works regardless of Seller Central's display language (e.g. "192 fulfillable" in English, "192 주문 처리 가능" in Korean)
- Auto-selects Standard shipping speed
- Works with Seller Central's Korean-language field labels (전체 이름, 상세 주소, 도시, 시/도, 우편번호, 전화번호, 이메일 주소, 국가) in addition to English
- Reads a Ctrl+V clipboard paste as a fallback autofill path when the URL-hash bridge from GCX Reply didn't carry data

---

### Amazon JP MCF Autofill (`v1.5.2`)
**Matches:** `sellercentral-japan.amazon.com` — MCF create-order pages

JP-specific variant of the MCF Autofill script. Picks up order data from GCX Reply's MCF button, maps Japanese prefecture names to their romanized equivalents, and autofills the JP MCF order form.

---

### Amazon Invoice Automation (`v1.5`)
**Matches:** `sellercentral.amazon.de` — individual order pages

Adds a "Run Now" button on Amazon.de Seller Central order pages. On click, attempts to download the deemed resale/supply invoice first, falling back to the Amazon-generated invoice. Copies the result to clipboard via `GM_setClipboard`.

---

### GChat Reply Suggest (`v3.7.0`)
**Matches:** `chat.google.com/*` and `spigenhelp.zendesk.com/agent/tickets/*` (works inside the Chrome-PWA "desktop app" install too, since it's the same Chrome origin/extension context)

Alt+G behaves differently depending on the Chat room. **The deterministic ticket-forward flow is the default in every room** — AI-suggested replies are now the exception, opted into per-room, not the other way around.

**In designated AI-suggest rooms** (`AI_SUGGEST_ROOM_IDS` near the top of the script, empty until a room is added) — suggests 3 short candidate reply sentences for the current conversation as clickable chips docked above the compose box (AI-generated). Click a chip to insert it at the cursor; Esc dismisses the bar.
- Grabs the tail (~2000 chars) of `[role="main"]`'s `innerText` as conversation context — deliberately avoids Google Chat's hashed/obfuscated class names, which rotate on redeploy
- Picks the last visible `[role="textbox"][contenteditable="true"]` on the page as the active compose box (Google Chat renders more than one; the others are hidden thread/history boxes)
- Calls a **local** backend (`../gchat_reply_suggest_server.py`, listens on `127.0.0.1:8765`) rather than a cloud API — the backend shells out to the existing `claude` CLI (`~/.local/bin/claude -p`), so it reuses the Claude Code subscription with no separate API key or per-token billing
- Backend must be running for suggestions to work: `python3 ~/Desktop/GCX/Browser_Extensions/gchat_reply_suggest_server.py` (not yet wired to auto-start on login — run manually, or ask to set up a LaunchAgent)

**Everywhere else (the default)** — Alt+G first shows a **T3 Esc / Gratitude / Reminder** picker. Rooms are opted OUT of this (into AI-suggest instead) by Chat **space ID** (the `/app/chat/<ID>` segment of the URL), not by display name — `document.title` turned out to be unreliable for this (Chat rewrites it for unread counts, "X messaged you", etc., not just a stable "RoomName - Chat" string). Configured in `AI_SUGGEST_ROOM_IDS` near the top of the script; **to add a room** (e.g. once the planned "GCX T2 ESC. Ticket 보고" room exists), open it, press Option+G once, then check the browser console — the script logs the current space ID and name every time, ready to copy into the config.

**All pickers** (ticket/mention/honorific/index) listen for ↑/↓/Enter/Escape on the **capture phase** with `stopImmediatePropagation()`, not just `preventDefault()`. The **mention picker specifically wraps** (`openPicker(..., { wrap: true })`): ↑ on the first option jumps to the last, ↓ on the last jumps back to the first — the other three pickers (ticket/honorific/index) still clamp at each end, since only the mention list was asked to wrap. This mattered in practice: `preventDefault()` alone only cancels the *browser's* default action for a key — it does nothing to stop *other* JS listeners (Chat's own) from also reacting to the same event. Chat treats ↑ in an empty compose box as "edit my last sent message" (a common chat-app convention), so navigating a picker with ↑ was silently also flipping the previous sent message into inline edit mode underneath it. Capture-phase + `stopImmediatePropagation()` means this script's handler runs and fully consumes the event before Chat's own bubbling-phase listener ever sees it, regardless of registration order (Chat's script loads well before this one, so on the bubbling phase alone it would always win the race).

**T3 Esc** — the deterministic, no-AI ticket-forward flow:
- **Shopify EU order number (v3.7.0):** the Zendesk side also records `Requested Channel type*` (field resolved by title from `ticket_fields.json`, since no ID is recorded) and `Order ID` (`360021934132`). When the channel is `Shopify EU` and Order ID is set, the T3 Esc draft gets ` (주문번호: #1763)` appended after "…감사하겠습니다!" (a leading `#` is normalized, never doubled). Tickets recorded before the update lack these fields — revisit them once to refresh.
- The Zendesk-side half of this same script (`@match spigenhelp.zendesk.com/agent/tickets/*`) records Country/Device/Product Name/ASIN (same custom-field IDs as GCX Reply's `ZD` constant) for every ticket you visit, into a shared list of the 10 most recent
  - **Zendesk is a single-page app** — clicking between tickets (a list, a "next ticket" button, etc.) changes the URL without a full page reload, so recording once when the script loads only ever captured whichever ticket happened to be open at that moment; anything visited afterward via in-app navigation silently never got recorded. Fixed by polling `location.pathname` for the ticket ID to change (every 1.5s) and re-recording each time, instead of a one-shot call. `recordCurrentTicket()` also now logs to the console on both success and failure (`[GChat Reply Suggest] recorded ticket #...`) for easy self-diagnosis if a ticket still doesn't show up in the picker
- Shows that list (most recent pre-highlighted); **↑/↓ arrow keys browse older tickets, Enter confirms** (mouse click also works). Picking one opens a second picker — **who to @mention** (defaults to `(no mention)`, plus every sender seen in the visible conversation) — then a third picker — **which honorific** (`리더`/`파트장`/`프로`, defaults to `프로` — configured in `HONORIFIC_TITLES`/`DEFAULT_HONORIFIC` near the top of the script) — then a fourth picker — **which index number** (see below) — then inserts:
  ```
  안녕하세요 {리더|파트장|프로}님, 담당하시는 제품 관련 (사유) 문의가 들어와 전달드립니다. 확인 후 회신해 주시면 감사하겠습니다!

  **{일련번호} {국가} {기종}용 {제품명} [{ASIN}]**
  {티켓 링크, 실제 하이퍼링크}
  ```
  The reference line renders **bold** and the ticket link is a real clickable hyperlink — both verified to survive an actual send, not just the draft view. (Chat's compose box enforces a Trusted Types CSP that blocks `execCommand("insertHTML", ...)`; formatting is applied by inserting plain text first, then wrapping the exact substrings with real DOM nodes — `document.createElement("b"/"a")` — which sidesteps Trusted Types entirely since no HTML string is ever parsed.)
  - **The `(사유)` bracket is only the part after the last `_`** in Zendesk's raw option label — those labels are usually `"{English category}_{Korean detail}"` (e.g. `"(Sales Inquiry)_재입고관련문의"`, `"(Urgent)_리스팅오류"`), and only the Korean detail (`재입고관련문의`, `리스팅오류`) belongs in the sentence — the English/bracketed category prefix is dropped (`extractReasonLabel()`). Falls back to the whole label if there's no underscore, and to `문의` if there's no label at all.
  - **The index number is a user pick, not an auto-scanned guess.** An earlier version tried scanning the room's actual messages for the highest reference-line index used and auto-inserting the next one; per explicit request this was replaced with an explicit **1–15 dropdown** (`MAX_DAILY_INDEX`) — 1 always at the top, Enter-with-no-navigation picks 1. It's pre-highlighted on "next after whatever was last confirmed sent today in this room" (e.g. if 1 was already sent, 2 is pre-highlighted) — but the full 1–15 list is still shown, so ↑ once from that default reaches 1 again in case you need to reuse it. The underlying counter (`peekNextIndexForToday()` read-only for the default; `commitIndexUsed()` the only place that persists) is keyed by room + day, so it resets automatically at midnight KST with no separate reset logic — and it only advances once a draft is confirmed **actually sent** (via the "drafted" → "sent" watcher transition), not the moment the picker completes, so re-opening the picker and picking a different ticket, or aborting and redrafting, doesn't burn index numbers that were never sent.
  - **@mention is embedded as plain text** (`@{name} `, prefixed to the message) — a deliberate choice, not a shortcut. A real, notification-triggering mention can only come from Chat's own People autocomplete, which requires a genuinely user-trusted keystroke, and this was verified two different ways: a synthetic JS keydown/beforeinput/keyup "@" sequence produced zero DOM changes anywhere on the page, and even a genuinely trusted **CDP-level** keystroke — the mechanism Playwright/Puppeteer use, `isTrusted: true`, indistinguishable from real hardware input at the browser's own trust layer — still didn't trigger Chat's popup. So this isn't just an `event.isTrusted` check to route around; there's no scriptable path to a real mention at all, only a genuine physical `@` keypress from the user.
- After a forward, the script watches for the next incoming message in that room and auto-offers the same @mention + honorific picker pair for a confirm-reply (reusing the same bold/hyperlinked ticket reference line), pre-selecting whoever just replied for the @mention step
  - **This watcher only starts once the forward is actually posted, not the moment it's drafted.** The compose box lives *inside* `[role="main"]` (verified: `main.contains(composeBox)` is true), so any DOM scan of the message area needs to explicitly exclude the compose box's subtree — otherwise the still-unsent draft looks indistinguishable from an already-sent message, and any unrelated message arriving while you're still editing the draft gets misread as "someone replied," popping the confirm-picker and appending onto your draft when you click something. Fixed with a `getMessageOnlyText()` helper (DOM Range that explicitly ends right before the compose box), plus a two-stage `"drafted" → "sent"` state that only flips once the forward's own reference line is found in the *message-only* text, with a 1.5s cooldown afterward before any further mutation counts as a possible reply
- The member list offered in the @mention picker comes from distinct sender names seen in the currently-visible conversation (not the space's official member list, which would need navigating Chat's own People panel) — good enough in practice, but someone who hasn't posted recently won't show up

**Gratitude** / **Reminder** — meant to be used from a **thread** reply box (a T3 Esc message's "N replies" thread opened via Chat's own thread panel), not the main room. Auto-reads the @mention name + honorific from the T3 Esc root message the open thread is attached to (no re-picking) and inserts:
  ```
  {이름} {honorific}님, 확인 감사합니다. 주신 메모 확인 후 처리하도록 하겠습니다!            ← Gratitude

  안녕하세요 {이름} {honorific}님. 해당 티켓에 대한 답변이 아직 없어서 업무 마감차 리마인드챗드립니다.
  시간 되실 때 확인 한 번 부탁드립니다. 감사합니다!                                    ← Reminder
  ```
  If no thread is open, or the thread's root wasn't a T3 Esc message from this script, it falls back to a name-less version with the default honorific (`프로님, 확인 감사합니다. ...` / `안녕하세요 프로님. ...`) rather than an error chip — per explicit request, a generic reply is still useful there.
  - **Finding "the thread's root message" turned out to be genuinely hard via DOM geometry** — the Thread side panel's exact container boundary doesn't stand out reliably from the main room panel (both live under a shared ancestor). The approach that actually worked (`getThreadRootInfo()`): search the whole page for our own script's distinctive bold reference-line signature, restricted to the **right half of the viewport** (verified live: the thread panel's message copies sit at a consistently larger `x` than the main panel's), and take the **topmost** match, since the root message is pinned above the reply list. Then parse `@{name}\n안녕하세요 {honorific}님` out of that message's text (handles both this script's own plain-text mention format and a real Chat-native mention chip's slightly different whitespace/newline rendering — verified against actual real messages in production use, not synthetic test data). This only works for threads whose root was actually sent via the T3 Esc flow (has the bold ref line) — a thread on an older, pre-formatting message won't be found, hence the error-chip fallback
    - **The signature match originally required a trailing `[ASIN]` bracket** (6–12 alphanumeric characters) — a ticket with no ASIN on file produces a malformed/empty bracket that silently never matched, so Gratitude/Reminder reported "couldn't find a T3 Esc message" even with the root message plainly visible. Fixed on both ends: `referenceLineText()` now omits the bracket entirely when there's no ASIN (rather than leaving an empty `[]`), and the detection regex no longer depends on the bracket at all — it anchors on the `용` particle (the device-purpose suffix, e.g. "iPhone 15용") that always sits between device and product name in this template, which is present regardless of ASIN
  - `getComposeBox()` (unchanged) already picks the thread's own "Reply" box correctly when a thread panel is open, since it renders after the main panel in DOM order — verified live, no changes needed there

---

## Installation

1. Install the [Tampermonkey extension](https://www.tampermonkey.net/) in Chrome.
2. Get the `.user.js` file for the script you want to install.
3. Drag the file into the Tampermonkey Dashboard, or open it and click "Install" when prompted.

Once installed, Tampermonkey checks the `@updateURL` in the script header and auto-updates when a new version is pushed.

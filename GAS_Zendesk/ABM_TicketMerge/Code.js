/********************************
 * ABM Ticket Merge
 * ------------------------------
 * Amazon sends an unthreaded "You have received a message" notification
 * email for every single ABM message, so Zendesk's mail channel creates a
 * brand new ticket per message instead of appending to the buyer's existing
 * open ticket (unlike Seller Central's own inbox, which threads by caseId).
 *
 * A Zendesk Trigger (condition: ticket created AND tags contain
 * "buyer_message_amazon") calls this Web App via a Notify-target webhook,
 * passing the new ticket's ID. If the same buyer already has a prior
 * buyer_message_amazon ticket, this script:
 *   1. Posts the new ticket's message as a public comment on that prior
 *      ("primary") ticket (authored as the requester, so it reads like the
 *      customer's own follow-up — matching how Seller Central threads it),
 *      REOPENING the primary if it had already been solved.
 *   2. Closes the new (duplicate) ticket with an internal note pointing to
 *      the ticket it was merged into.
 * If no reusable prior ticket exists, the new ticket is left untouched — it
 * becomes the thread's primary ticket for any future follow-ups.
 *
 * Primary selection (v2 — reopen & merge, case-ID aware):
 *   - Candidates = same-requester, buyer_message_amazon-tagged tickets that
 *     are NOT closed (Zendesk 'closed' is terminal — can't be reopened or
 *     commented on), newest first.
 *   - Amazon's buyer proxy address embeds the Seller Central case ID
 *     (e.g. "...+76c079d3-...@marketplace.amazon.co.jp"). When the NEW ticket
 *     has one, prefer the newest candidate with the SAME case ID — that's the
 *     exact key Seller Central threads by, and it avoids wrongly merging two
 *     genuinely different open cases from the same buyer. Candidates whose
 *     address has no resolvable case ID (older tickets, empty from-address)
 *     are eligible as a soft fallback; a candidate with a DIFFERENT explicit
 *     case ID is never chosen.
 *   - Earlier behavior only matched still-OPEN tickets and never reopened,
 *     so once an agent solved a ticket, the buyer's next message spawned a
 *     fresh unmerged ticket. That was the root cause of the "New tickets keep
 *     getting created" reports (e.g. JP chain #1000153447→603→740, all solved
 *     between messages).
 ********************************/

const ZENDESK_EMAIL = 'kjw@spigen.com';
const ZENDESK_TOKEN = 'QhM2AiBYwTZTSb04Qjor918PHtttxp8xAzCFfFsg';
const ZENDESK_SUBDOMAIN = 'spigenhelp';
const ABM_TAG = 'buyer_message_amazon';

// Shared secret the Zendesk webhook target must send back (see setupInstructions_ below).
// Checked against the `?secret=` query param on the webhook URL.
const WEBHOOK_SECRET = 'xN61CnX8OWX1O3lquwj-JFi6YTmFeezTaVhHsWZXvi8';

function zdAuthHeader_() {
  return 'Basic ' + Utilities.base64Encode(`${ZENDESK_EMAIL}/token:${ZENDESK_TOKEN}`);
}

// Found live 2026-07-30 investigating ticket #1000155549 (never got its
// inbound cleanup/collapse pair): the webhook trigger fired and returned
// HTTP 200, but handleNewAbmTicket_ had actually thrown mid-run —
// `Exception: Address unavailable: https://spigenhelp.zendesk.com/api/v2/...`
// — a transient GAS UrlFetchApp connection-level failure that `muteHttpExceptions`
// does NOT catch (that flag only suppresses non-2xx HTTP responses; this is a
// failure to even get a response). Since Apps Script's ContentService can't
// return a real non-200 status from doPost, Zendesk had no signal to retry —
// it saw 200 and considered the webhook successfully delivered, permanently
// stranding the ticket. Scanned every webhook invocation over the past week:
// found 4 tickets hit this exact pattern (on 3 different endpoints — getUser_,
// search.json, getTicket_ — confirming it's generic UrlFetchApp flakiness, not
// one bad call site), roughly one every few days. Fixed at the shared fetch
// layer so every call site gets the retry, not just the ones seen so far.
const ZD_FETCH_RETRIES = 3;
function zdFetch_(path, options) {
  const url = path.startsWith('http') ? path : `https://${ZENDESK_SUBDOMAIN}.zendesk.com${path}`;
  const opts = Object.assign({
    headers: { Authorization: zdAuthHeader_() },
    contentType: 'application/json',
    muteHttpExceptions: true
  }, options || {});
  for (let attempt = 1; attempt <= ZD_FETCH_RETRIES; attempt++) {
    let res;
    try {
      res = UrlFetchApp.fetch(url, opts);
    } catch (e) {
      // A Workspace-account-wide daily UrlFetchApp quota exhaustion (seen
      // live 2026-08-18: "하루에 premium urlfetch 서비스를 너무 많이
      // 호출했습니다") throws from UrlFetchApp.fetch() exactly like a
      // transient connection failure, so the retry loop below used to burn
      // all 3 attempts (with growing sleeps) on every single call anyway —
      // pointless, since this can't possibly succeed again until the quota
      // resets, and it eats into GAS's 6-minute execution ceiling for every
      // subsequent zdFetch_ call in the same run. Fail immediately instead.
      if (/urlfetch/i.test(e.message) && /하루|quota|exceeded/i.test(e.message)) {
        throw new Error(`Zendesk API ${opts.method || 'GET'} ${url} -> daily UrlFetchApp quota exhausted (not retrying, won't recover until reset): ${e.message}`);
      }
      if (attempt < ZD_FETCH_RETRIES) { Utilities.sleep(500 * attempt); continue; }
      throw new Error(`Zendesk API ${opts.method || 'GET'} ${url} -> connection failed after ${ZD_FETCH_RETRIES} attempts: ${e.message}`);
    }
    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code >= 300) {
      throw new Error(`Zendesk API ${opts.method || 'GET'} ${url} -> ${code}: ${body}`);
    }
    return body ? JSON.parse(body) : {};
  }
}

function getTicket_(ticketId) {
  return zdFetch_(`/api/v2/tickets/${ticketId}.json`).ticket;
}

function getUser_(userId) {
  return zdFetch_(`/api/v2/users/${userId}.json`).user;
}

// First public comment on a ticket (the actual Amazon notification body).
function getFirstComment_(ticketId) {
  const data = zdFetch_(`/api/v2/tickets/${ticketId}/comments.json?sort_order=asc`);
  const comments = data.comments || [];
  return comments.length ? comments[0] : null;
}

// Re-hosts a source comment's attachments onto Zendesk so they can be attached
// to the primary ticket's comment. A Zendesk comment can't reference another
// ticket's attachment tokens directly — each file must be downloaded and
// re-uploaded to get a fresh upload token. Returns an array of upload tokens
// (one per successfully transferred file) to pass as `comment.uploads`.
// Best-effort per file: a single download/upload failure is logged and skipped
// so the rest of the message (text + other attachments) still merges.
function transferAttachments_(attachments) {
  const tokens = [];
  (attachments || []).forEach(att => {
    try {
      // content_url carries its own access token, but send the API Basic auth
      // too so private/agent-only attachments download reliably.
      const dl = UrlFetchApp.fetch(att.content_url, {
        headers: { Authorization: zdAuthHeader_() },
        muteHttpExceptions: true
      });
      if (dl.getResponseCode() >= 300) {
        Logger.log(`transferAttachments_: download failed for ${att.file_name} (${dl.getResponseCode()})`);
        return;
      }
      const blob = dl.getBlob().setName(att.file_name);
      const up = UrlFetchApp.fetch(
        `https://${ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/uploads.json?filename=${encodeURIComponent(att.file_name)}`,
        {
          method: 'post',
          contentType: att.content_type || 'application/octet-stream',
          payload: blob.getBytes(),
          headers: { Authorization: zdAuthHeader_() },
          muteHttpExceptions: true
        }
      );
      if (up.getResponseCode() >= 300) {
        Logger.log(`transferAttachments_: upload failed for ${att.file_name} (${up.getResponseCode()}): ${up.getContentText()}`);
        return;
      }
      const token = JSON.parse(up.getContentText())?.upload?.token;
      if (token) tokens.push(token);
    } catch (err) {
      Logger.log(`transferAttachments_: error on ${att && att.file_name} — ${err}`);
    }
  });
  return tokens;
}

/********************************
 * ABM inbound cleanup
 * ------------------------------
 * Amazon's "You have received a message" notification email buries the
 * buyer's own typed text inside a full marketing/legal HTML template (logo,
 * order table, "did this email help" survey buttons, footer copyright,
 * opt-out text, commMgrTok/SPC-xxAmazon tracking IDs) — the ticket comment
 * Zendesk creates from it looks nothing like a normal claim, even though
 * Zendesk's own inbound mail parsing already attaches the buyer's real
 * photos/PDFs correctly (verified live on ticket #1000154136 — no fix
 * needed there).
 *
 * Investigated live 2026-07-21 by diffing the raw html_body of several real
 * ABM tickets across JP/EN/DE: every sample wraps the buyer's own text in
 * exactly one <pre> block with an IDENTICAL inline style across every
 * marketplace/language sampled — only the surrounding template strings are
 * translated, this wrapper is not. That makes it a reliable, locale-agnostic
 * extraction anchor, so no Seller Central lookup (and therefore no live
 * agent browser session) is needed at all — this runs entirely server-side.
 *
 * Zendesk's "Ticket Comments" API has no way to replace a comment's visible
 * content with new text, so this posts a SEPARATE clean public comment
 * authored as the requester (same technique already proven in
 * mergeNewTicketIntoPrimary_) with the buyer's real attachments re-hosted
 * onto it. The original raw comment is left fully intact — zero data loss,
 * fully auditable, nothing destructive.
 *
 * (Briefly tried also redacting the raw original's text, GAS v14-v18,
 * 2026-07-22 — reverted the same day after discovering Zendesk's own native
 * "merge tickets" feature quotes a merged ticket's last comment by its
 * CURRENT stored text, so an agent manually merging a duplicate ABM ticket
 * whose raw comment had already been redacted produced a merge-quote system
 * note full of garbled block characters instead of the real message.
 * Redaction is permanent/global — it doesn't just affect this feature's own
 * view of the comment. Not worth that risk; see the removed
 * redactCommentText_ function's final comment and [[abm_ticket_merge]]
 * memory for the full incident writeup.)
 ********************************/

// Same inline style Amazon's template applies to the buyer's own message
// text, confirmed identical (byte-for-byte on this substring) across every
// JP/EN/DE sample checked — a couple of samples drop the single leading
// space before "color: black" (seemingly whenever the message body is just
// a bare URL), hence \s* rather than a literal space.
const ABM_RAW_PRE_RE = /<pre[^>]*white-space:\s*-o-pre-wrap[^>]*>([\s\S]*?)<\/pre>/i;

function decodeHtmlEntities_(s) {
  return String(s || '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCharCode(parseInt(n, 16)))
    .replace(/&amp;/gi, '&'); // last: avoid double-decoding entities that start with &amp;
}

// Extracts just the buyer's own typed text from a raw ABM notification
// email's html_body. Returns:
//   null       — the <pre> marker wasn't found at all (not a raw ABM
//                template — e.g. already cleaned, or a shape we haven't
//                seen) — callers must NOT touch the comment in this case.
//   ''         — marker found but empty (e.g. a photo-only message with no
//                typed text) — callers should still surface a clean copy so
//                the real attachments aren't stranded inside the raw wall
//                of HTML, just with a placeholder line instead of blank text.
//   non-empty  — the buyer's real message, tags stripped, entities decoded.
function cleanAbmMessageText_(html) {
  const m = ABM_RAW_PRE_RE.exec(html || '');
  if (!m) return null;
  const stripped = m[1].replace(/<[^>]+>/g, '');
  return decodeHtmlEntities_(stripped).trim();
}

// Idempotency tag for a cleaned source comment. Previously this tracking
// marker was embedded as a visible first line of the public comment itself
// ("[Amazon Buyer Message — auto-cleaned copy of comment N]"), which showed
// up at the top of every cleaned message in the ticket's own conversation
// thread — moved to a tag instead so the public comment is just the buyer's
// real message, nothing else.
function abmCleanedTag_(sourceCommentId) {
  return `abm_cleaned_${sourceCommentId}`;
}

// Adds a tag to a ticket WITHOUT replacing its existing tags. The single-
// ticket update endpoint (PUT /api/v2/tickets/{id}.json) has no
// `additional_tags` field — only a plain `tags` field that REPLACES the
// whole array (confirmed via Zendesk's own API docs after the first version
// of this fix silently no-opped: passing `additional_tags` alongside
// `comment` in that same PUT was just an unrecognized/ignored field, so NO
// tag was ever actually added in production, and every subsequent cleanup
// run had zero idempotency signal — reproduced live on ticket #1000153636,
// where two ABM messages 23 seconds apart each triggered their own
// cleanupExistingAbmTicket_ run, both saw the first message as un-cleaned,
// and both posted a duplicate clean copy of it). Adding tags without
// clobbering existing ones requires the DEDICATED tags endpoint instead:
// PUT /api/v2/tickets/{id}/tags.json with body {tags: [...]}.
function addTicketTag_(ticketId, tag) {
  zdFetch_(`/api/v2/tickets/${ticketId}/tags.json`, {
    method: 'put',
    payload: JSON.stringify({ tags: [tag] })
  });
}

// REVERTED 2026-07-23 (was live briefly as redactCommentText_, GAS v14-v18):
// tried auto-redacting the raw comment's text right after posting its clean
// copy, so only the clean message would be visible. Discovered a real,
// unrelated side effect that made this not worth the risk: Zendesk's own
// NATIVE "merge tickets" feature (an agent manually merging a duplicate ABM
// ticket via the Zendesk UI — a completely different mechanism from this
// project's own auto-merge) quotes the merged ticket's "last comment" by
// reading whatever its CURRENTLY STORED text is. On ticket #1000154651, an
// agent natively merged in ticket #1000154652 hours after its raw comment
// had already been redacted — the resulting merge-quote system note showed
// the redacted block-character garbage verbatim, not the original message.
// Redaction is permanent/global: it doesn't just affect how a comment looks
// in its own ticket, it destroys the text for ANY other Zendesk feature that
// might read it later (native merges being the one that surfaced first).
// User's call, given that risk: keep the original (pre-v14) design — post
// the clean copy ALONGSIDE the untouched raw comment, never redact.
// Tickets processed while v14-v18 was live (roughly 2026-07-22 08:00 through
// 2026-07-23) have their raw comments PERMANENTLY blanked already — Zendesk
// redaction cannot be undone via the API. That set of affected real tickets
// was small (this feature was live under a day); see [[abm_ticket_merge]]
// memory for the full incident writeup and any tickets identified.

function alreadyCleanedSourceIds_(ticket, comments) {
  const ids = new Set();
  const tagRe = /^abm_cleaned_(\d+)$/;
  (ticket.tags || []).forEach(t => {
    const m = tagRe.exec(t);
    if (m) ids.add(m[1]);
  });
  // Back-compat: tickets cleaned before this tag existed still carry the old
  // visible "[Amazon Buyer Message — auto-cleaned copy of comment N]" marker
  // in their clean comment's body instead of a tag — keep recognizing those
  // too so this migration doesn't re-post duplicates on tickets cleaned under
  // the old scheme. New cleanups only ever add the tag, never this text.
  const legacyRe = /auto-cleaned copy of comment (\d+)/g;
  (comments || []).forEach(c => {
    const text = c.body || c.plain_body || '';
    let m;
    while ((m = legacyRe.exec(text)) !== null) ids.add(m[1]);
  });
  return ids;
}

// ── Auto-fill ticket fields from the order GCX Reply's Auto-Fill would resolve ──
// User's ask: Order ID / Customer Full Name / Amazon Fulfillment Methods /
// Country should already be filled in by the time an agent opens a new ABM
// ticket, not require them to click GCX Reply's Auto-Fill button first — only
// the buyer's actual message should read as "what was received."
//
// Field IDs must match tampermonkey_scripts/GCX Reply.user.js's ZD constant
// exactly (same Zendesk instance, same custom fields).
const ZD_ORDER_ID    = 360021934132;
const ZD_CUST_NAME   = 360021999951;
const ZD_COUNTRY     = 4513936822297;
const ZD_FULFILLMENT = 900002781823;
const ZD_ASIN        = 360021934312;

// Mirrors GCX Reply's COUNTRY_MAP exactly — must match, since these are the
// literal Zendesk tagger field option values (confirmed live against
// ticket_fields/4513936822297.json: de/uk/fr/it/es/nl/se/ie/pl/tr/be/in/jp/
// sg/au/us/ca/mx/kr).
const ABM_COUNTRY_MAP = {
  US: 'us', GB: 'uk', DE: 'de', FR: 'fr', IT: 'it', ES: 'es', JP: 'jp',
  NL: 'nl', SE: 'se', IE: 'ie', PL: 'pl', TR: 'tr', BE: 'be', IN: 'in',
  SG: 'sg', AU: 'au', CA: 'ca', MX: 'mx', KR: 'kr',
};

// Amazon buyer-proxy address domain → the same Country field value strings
// as ABM_COUNTRY_MAP above (e.g. "...+<uuid>@marketplace.amazon.co.uk" →
// 'uk'). Lets Country get filled straight from the inbound ABM message
// itself, with NO dependency on a resolved order — unlike the SP-API
// address lookup below, this also covers order-less ABM tickets (a
// pre-purchase question with no Order ID anywhere in the message, e.g.
// #1000154804), which previously left Country blank forever since there
// was no order to resolve it from.
//
// de/uk/be/es/fr/in/it/jp confirmed live by cross-referencing 100 real ABM
// tickets' via.source.from.address domain against their ALREADY-correct
// (order-resolved) Country field value. The rest (nl/se/ie/pl/tr/sg/au/us/
// ca/mx/kr) are standard/well-known Amazon marketplace domains but did not
// appear in that sample — unverified against a real ticket, so double-check
// if one of these ever looks wrong live.
const ABM_MARKETPLACE_DOMAIN_COUNTRY_ = {
  'amazon.de': 'de',            // confirmed
  'amazon.co.uk': 'uk',         // confirmed
  'amazon.fr': 'fr',            // confirmed
  'amazon.it': 'it',            // confirmed
  'amazon.es': 'es',            // confirmed
  'amazon.co.jp': 'jp',         // confirmed
  'amazon.in': 'in',            // confirmed
  'amazon.com.be': 'be',        // confirmed
  'amazon.nl': 'nl',            // unverified
  'amazon.se': 'se',            // unverified
  'amazon.ie': 'ie',            // unverified
  'amazon.pl': 'pl',            // unverified
  'amazon.com.tr': 'tr',        // unverified
  'amazon.sg': 'sg',            // unverified
  'amazon.com.au': 'au',        // unverified
  'amazon.com': 'us',           // unverified
  'amazon.ca': 'ca',            // unverified
  'amazon.com.mx': 'mx',        // unverified
  'amazon.co.kr': 'kr',         // unverified
};

// Same address caseIdFromTicket_ reads (the Amazon buyer-proxy from-address,
// e.g. "s0574jllj4n84kf+<uuid>@marketplace.amazon.co.jp") — pulls out the
// "amazon.<tld>" domain instead of the case-id UUID.
function countryFromTicketMarketplace_(ticket) {
  const addr = (ticket && ticket.via && ticket.via.source && ticket.via.source.from
    && ticket.via.source.from.address) || '';
  const m = addr.match(/@[\w-]*\.?(amazon\.[\w.-]+)$/i);
  const domain = m ? m[1].toLowerCase() : null;
  return domain ? (ABM_MARKETPLACE_DOMAIN_COUNTRY_[domain] || null) : null;
}

// Rejects a candidate "name" that's actually just the buyer proxy address's
// local part (e.g. "pzvz36hxbwx0l9z" from
// "pzvz36hxbwx0l9z+18bbf0e8-...@marketplace.amazon.fr"), title-cased or not.
// Found live on ticket #1000154946: SP-API's own BuyerName field returned
// this exact proxy string (Title-cased) for a privacy-masked EU order — with
// no validation, that got written straight into Customer Full Name instead of
// the real name ("Julien") that was sitting right there in
// ticket.via.source.from.name the whole time. An agent had to manually
// correct it. Guards BOTH candidate sources (SP-API BuyerName and the
// Zendesk from-name) since either could theoretically be proxy-shaped.
function looksLikeAbmProxyName_(name, ticket) {
  if (!name) return false;
  const addr = (ticket && ticket.via && ticket.via.source && ticket.via.source.from
    && ticket.via.source.from.address) || '';
  const localPart = addr.split('@')[0].split('+')[0];
  if (!localPart) return false;
  return name.replace(/[^a-z0-9]/gi, '').toLowerCase() === localPart.replace(/[^a-z0-9]/gi, '').toLowerCase();
}

// Calls GCX Reply's OWN order-lookup endpoint (GCXReply_GAS's `?orderId=`,
// same SP-API-backed data GCX Reply's Auto-Fill button uses) rather than
// re-implementing SP-API SigV4 signing in this project too.
function fetchGcxOrderData_(orderId) {
  const res = UrlFetchApp.fetch(`${GCX_GAS_URL}?orderId=${encodeURIComponent(orderId)}`,
    { muteHttpExceptions: true, followRedirects: true });
  try {
    const d = JSON.parse(res.getContentText());
    return d && !d.error ? d : null;
  } catch (_) { return null; }
}

// Same precedence GCX Reply's own Auto-Fill uses (buildAndShow's
// fulfillmentReady): a PAN EU seller SKU suffix wins regardless of
// FulfillmentChannel; otherwise AFN→fba, MFN→merchant__fbm_. Confirmed live
// against ticket_fields/900002781823.json for the exact tag values.
function abmFulfillmentTag_(fulfillmentChannel, sellerSku) {
  if (/pan|eup/i.test(sellerSku || '')) return 'pan_eu';
  if (fulfillmentChannel === 'AFN') return 'fba';
  if (fulfillmentChannel === 'MFN') return 'merchant__fbm_';
  return null;
}

// Fills Order ID / Customer Full Name / Country / Amazon Fulfillment Methods /
// ASIN from the order GCX Reply's own Auto-Fill would resolve — but ONLY
// fields that are CURRENTLY EMPTY, so this never overwrites a value an agent
// (or an earlier run) already set. Safe/idempotent to call on every cleanup
// run; no-ops instantly (no HTTP calls) once all 5 fields are already
// populated.
//
// Order ID / Country / Fulfillment genuinely need a resolved order (SP-API
// lookup via GCX Reply's own endpoint) — but Customer Full Name and ASIN
// don't: ASIN is embedded directly in the raw ABM email regardless of
// whether the message is about an actual order, and Customer Full Name
// falls back to the ticket requester's from-name. The original version
// gated on `if (!orderId) return`, which — for an order-less ABM ticket
// (e.g. a pre-purchase compatibility question with no Order ID anywhere in
// the message, confirmed live on ticket #1000154672) — skipped ALL fields,
// including the two that never needed an order in the first place.
function autoFillAbmTicketFields_(ticket, orderId, asin) {
  const cf = ticket.custom_fields || [];
  const valueOf_ = id => { const f = cf.find(x => x.id === id); return f ? f.value : null; };
  const needs = {
    orderId:     !valueOf_(ZD_ORDER_ID),
    custName:    !valueOf_(ZD_CUST_NAME),
    country:     !valueOf_(ZD_COUNTRY),
    fulfillment: !valueOf_(ZD_FULFILLMENT),
    asin:        !valueOf_(ZD_ASIN),
  };
  if (!needs.orderId && !needs.custName && !needs.country && !needs.fulfillment && !needs.asin) return;

  const custom_fields = [];

  if (needs.asin && asin) custom_fields.push({ id: ZD_ASIN, value: asin });

  // Resolved before the order fetch below so a country the marketplace
  // domain already answers doesn't force an otherwise-unneeded SP-API
  // round-trip (the common case now that most tickets carry a recognized
  // domain — see countryFromTicketMarketplace_).
  const domainCountry = needs.country ? countryFromTicketMarketplace_(ticket) : null;

  // Only worth the HTTP round-trip if orderId exists AND at least one
  // order-dependent field still needs filling.
  let data = null;
  if (orderId && (needs.orderId || needs.custName || (needs.country && !domainCountry) || needs.fulfillment)) {
    data = fetchGcxOrderData_(orderId);
  }

  if (needs.orderId && orderId) custom_fields.push({ id: ZD_ORDER_ID, value: orderId });
  if (needs.custName) {
    // Falls back to the requester's name (what Zendesk parsed from the ABM
    // email's From-name) when SP-API has no BuyerName, or when there was no
    // order to look up at all — same fallback ladder GCX Reply's own panel
    // uses. Either candidate can come back proxy-shaped (confirmed live for
    // SP-API's BuyerName on a privacy-masked EU order — see
    // looksLikeAbmProxyName_), so skip any candidate that's really just the
    // buyer proxy address instead of a name.
    const candidates = [
      (data && data.buyer && data.buyer.BuyerName) || null,
      (ticket.via && ticket.via.source && ticket.via.source.from && ticket.via.source.from.name) || null,
    ];
    const name = candidates.find(n => n && !looksLikeAbmProxyName_(n, ticket)) || null;
    if (name) custom_fields.push({ id: ZD_CUST_NAME, value: name });
  }
  if (needs.country) {
    // Marketplace domain first — works even with no order at all (fixes
    // #1000154804-style order-less ABM tickets, which previously left
    // Country blank forever since only the SP-API path below could fill
    // it). Falls back to the resolved order's shipping address only if the
    // domain wasn't recognized (see ABM_MARKETPLACE_DOMAIN_COUNTRY_'s
    // unverified entries).
    const countryVal = domainCountry
      || (data && ABM_COUNTRY_MAP[data.address && data.address.CountryCode]);
    if (countryVal) custom_fields.push({ id: ZD_COUNTRY, value: countryVal });
  }
  if (needs.fulfillment && data) {
    const sellerSku = data.items && data.items[0] ? data.items[0].SellerSKU : '';
    const fv = abmFulfillmentTag_(data.order && data.order.FulfillmentChannel, sellerSku);
    if (fv) custom_fields.push({ id: ZD_FULFILLMENT, value: fv });
  }
  if (!custom_fields.length) return;

  // Confirmed live (2026-07-22): a partial `custom_fields` array in a ticket
  // PUT updates ONLY the listed field IDs — every other custom field on the
  // ticket is left untouched (unlike `tags`, which replaces the whole array;
  // see addTicketTag_'s history of getting this exact distinction wrong for
  // a different field).
  zdFetch_(`/api/v2/tickets/${ticket.id}.json`, {
    method: 'put',
    payload: JSON.stringify({ ticket: { custom_fields } })
  });
}

// Scans a ticket for public, requester-authored comments that are still raw
// ABM template and don't already have a clean copy, and adds one clean
// public comment per such source comment — buyer's real text + re-hosted
// real attachments, authored as the requester so it reads as their own
// message. Fully idempotent (safe to re-run/backfill): each source comment
// gets at most one clean copy, tracked via abmCleanedTag_/
// alreadyCleanedSourceIds_ above. Covers both a ticket's own first
// (creation-time) comment and any raw comments mergeNewTicketIntoPrimary_
// posts onto an existing primary — same function handles both call sites.
// Also auto-fills Order ID/Customer Full Name/Country/Amazon Fulfillment
// Methods (see autoFillAbmTicketFields_) — only empty fields, every run.
//
// LockService-guarded: two ABM messages arriving seconds apart each trigger
// their own call to this function on the SAME ticket (handleNewAbmTicket_ /
// mergeNewTicketIntoPrimary_ per message) — without a lock, both could read
// the same "not yet cleaned" state before either had written its tag,
// racing into a duplicate clean copy of the same source comment (exactly
// what happened live on #1000153636 before this fix).
// onlyCommentId (optional): when set, only that source comment is eligible
// for cleaning — everything else on the ticket is left alone, even if it's
// still un-cleaned. Used by the live merge path (handleNewAbmTicket_) so one
// new message doesn't drag along a backfill of any OTHER old un-cleaned
// comment already sitting on the ticket, which reads as the whole thread
// re-arriving. Omit it (the manual/backfill call sites do) to scan and clean
// every un-cleaned comment on the ticket, same as before.
function cleanupExistingAbmTicket_(ticketId, onlyCommentId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { status: 'locked', ticketId };
  try {
    const ticket = getTicket_(ticketId);
    if (!ticket) return { status: 'not_found', ticketId };
    const requester = getUser_(ticket.requester_id);
    if (!requester) return { status: 'no_requester', ticketId };

    // ticket.description is Zendesk's own copy of the ticket's originating
    // comment's plain text — for an ABM ticket that's the raw "You have
    // received a message" email, which always embeds the ASIN and, when the
    // message concerns an actual order, the Order ID too (confirmed live on
    // #1000154560: description contains "402-1366488-5457149" verbatim; a
    // pre-purchase question like #1000154672 has an ASIN but no Order ID at
    // all). Same Order ID regex handleNewAbmTicket_/relayAbmReply_ use
    // elsewhere; ASIN regex matches GCX Reply's own `/\b(B[A-Z0-9]{9})\b/`.
    const orderIdMatch = (ticket.description || '').match(/\b(\d{3}-\d{7}-\d{7})\b/);
    const asinMatch = (ticket.description || '').match(/\b(B[A-Z0-9]{9})\b/);
    autoFillAbmTicketFields_(ticket, orderIdMatch ? orderIdMatch[1] : null, asinMatch ? asinMatch[1] : null);

    const data = zdFetch_(`/api/v2/tickets/${ticketId}/comments.json?include=users&sort_order=asc`);
    const comments = data.comments || [];
    // Identify "this is the buyer's own message" by ROLE, not by comparing
    // c.author_id to the CURRENT ticket.requester_id. Those can legitimately
    // diverge: normalizeAbmRequesterIdentity_ may have just merged this
    // ticket's original (proxy-only) end-user into a different canonical
    // end-user earlier in this SAME webhook invocation — Zendesk reassigns
    // the TICKET's requester_id on merge but does NOT retroactively rewrite
    // already-posted COMMENTS' author_id to match, so the ticket's own first
    // comment keeps pointing at the now-merged-away user id forever. An
    // author_id === requester.id check then silently skips it. Confirmed live
    // on #1000155389 and #1000155360 (2026-07-28, the day identity
    // normalization shipped): both ended up with cleanup: {cleaned: []} —
    // zero comments cleaned, so no clean copy ever got posted, which in turn
    // is exactly why GCX Reply's client-side raw-comment collapse never
    // fired on them (by design, it never collapses a raw comment with no
    // clean sibling to pair with — it wasn't a client-side bug at all, this
    // was the upstream cause).
    const roleById = {};
    (data.users || []).forEach(u => { roleById[u.id] = u.role; });
    const isBuyerPublicComment_ = c => c.public && roleById[c.author_id] !== 'agent' && roleById[c.author_id] !== 'admin';

    const done = alreadyCleanedSourceIds_(ticket, comments);
    // Defense in depth, independent of tag/marker bookkeeping: if a public,
    // buyer-authored comment with this EXACT clean text already exists,
    // treat it as already cleaned regardless — this is what actually caught
    // (and would have prevented) the #1000153636 duplicates, since the tag
    // write was silently failing at the time. Trade-off: a buyer who
    // legitimately repeats the exact same short message verbatim would only
    // get one clean copy — acceptable, this is a safety net, not the
    // primary mechanism (the tag is).
    const existingBodies = new Set(
      comments.filter(isBuyerPublicComment_)
        .map(c => (c.plain_body || c.body || '').trim())
    );

    // Posting a NEW public, buyer-authored comment onto a ticket that's
    // currently solved/pending/hold makes it look exactly like the customer
    // just replied — mergeNewTicketIntoPrimary_ already reopens the ticket in
    // the SAME PUT that posts a genuine live merge for exactly this reason,
    // but this function (used for backfilling old un-cleaned tickets) never
    // did. Confirmed live and caused real confusion on #1000155360
    // (2026-07-29): backfilling cleanup on an already-solved ticket posted
    // the buyer's clean-copy comment with no status change, so the ticket sat
    // "solved" despite a fresh public end-user comment. Zendesk's own native
    // reopen-on-end-user-comment automation eventually caught it ~5.5h later,
    // by which point it read as a mysterious duplicate customer message and
    // an agent had to manually investigate and re-solve it as NRN. Computed
    // once from the ticket's status at the START of this run — every posted
    // comment in this run reopens it the same way; harmless if it's already
    // open by a later comment in the same run.
    const shouldReopen = ticket.status === 'solved' || ticket.status === 'pending' || ticket.status === 'hold';

    const cleaned = [];
    comments.forEach(c => {
      if (onlyCommentId != null && String(c.id) !== String(onlyCommentId)) return;
      if (!isBuyerPublicComment_(c)) return;
      if (done.has(String(c.id))) return;
      const cleanText = cleanAbmMessageText_(c.html_body || '');
      if (cleanText === null) return; // not a raw ABM template comment — leave untouched

      const uploadTokens = transferAttachments_(c.attachments);
      const hasContent = cleanText || uploadTokens.length;
      if (!hasContent) return; // empty message, no attachments — nothing worth surfacing

      const bodyText = cleanText || '(메시지 텍스트 없음 — 첨부파일 참고)';
      if (existingBodies.has(bodyText.trim())) {
        // A clean copy already exists (created before the tag mechanism was
        // fixed) — just backfill the tag, don't post another one.
        addTicketTag_(ticketId, abmCleanedTag_(c.id));
        return;
      }

      const comment = { body: bodyText, public: true, author_id: requester.id };
      if (uploadTokens.length) comment.uploads = uploadTokens;
      const payload = { comment };
      if (shouldReopen) payload.status = 'open';

      zdFetch_(`/api/v2/tickets/${ticketId}.json`, {
        method: 'put',
        payload: JSON.stringify({ ticket: payload })
      });
      addTicketTag_(ticketId, abmCleanedTag_(c.id));
      existingBodies.add(bodyText.trim());
      cleaned.push({ sourceCommentId: c.id, reopened: shouldReopen, attachmentsTransferred: uploadTokens.length, attachmentsTotal: (c.attachments || []).length });
    });
    return { status: 'ok', ticketId, cleaned };
  } finally {
    lock.releaseLock();
  }
}

// Manual runner — run from the Apps Script editor (bypasses the webhook/
// secret entirely; calls the same logic directly). Used for backfilling
// tickets created before this feature existed.
function testCleanupOnTicket(ticketId) {
  const result = cleanupExistingAbmTicket_(ticketId);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

const CASE_ID_RE = /\+([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})@marketplace\./i;

function caseIdFromAddress_(addr) {
  const m = String(addr || '').match(CASE_ID_RE);
  return m ? m[1] : null;
}

// Seller Central case ID for a ticket. Confirmed live (2026-07-30, ticket
// #1000155523 investigation) that Amazon's ABM notification email NEVER
// carries the case-specific "+uuid" address in the literal From: header —
// `ticket.via.source.from.address` is always the bare buyer-proxy address
// (e.g. "wplt48syb1tpfwx@marketplace.amazon.de"). The "+uuid" address is only
// present as one of the message's OTHER original recipients (confirmed
// identical across an order-less ticket and an order-linked invoice-request
// ticket, so this isn't specific to one ABM message type) — Zendesk exposes
// that only via the ticket's first audit event
// (`via.source.from.original_recipients`), never on the plain ticket object.
// This means `.address` alone NEVER resolves a case ID in practice — every
// prior "confirmed working" case-ID resolution actually only ever worked via
// the caller's separate Order-ID/ASIN fallback (see relayAbmReply_ in GCX
// Reply.user.js). reconcileAbmRelays_ below has NO such fallback (it can't —
// order/ASIN matching needs a live, cookie-authenticated Seller Central
// browser session, which a GAS backend can't hold), so it silently could
// never rescue anything until this fix. Kept the from.address check FIRST
// (free, no extra API call) in case Amazon ever changes the email format to
// put it there directly; falls back to the audits lookup (one extra API
// call, only paid when the fast path misses).
function caseIdFromTicket_(ticket) {
  const addr = (ticket && ticket.via && ticket.via.source && ticket.via.source.from
    && ticket.via.source.from.address) || '';
  const direct = caseIdFromAddress_(addr);
  if (direct) return direct;
  return (ticket && ticket.id) ? caseIdFromOriginalRecipients_(ticket.id) : null;
}

function caseIdFromOriginalRecipients_(ticketId) {
  try {
    const data = zdFetch_(`/api/v2/tickets/${ticketId}/audits.json?sort_order=asc`);
    const first = (data.audits || [])[0];
    const recipients = (first && first.via && first.via.source && first.via.source.from
      && first.via.source.from.original_recipients) || [];
    for (const addr of recipients) {
      const c = caseIdFromAddress_(addr);
      if (c) return c;
    }
  } catch (e) {
    Logger.log(`caseIdFromOriginalRecipients_(${ticketId}): ${e.message}`);
  }
  return null;
}

// Amazon's ABM buyer-proxy address is unique PER CASE (the "+<uuid>" segment
// is the Seller Central case id — see caseIdFromTicket_ above), but the LOCAL
// PART BEFORE THE "+" is stable per real buyer on a given marketplace.
// Confirmed against 2 real examples: buyer "gg26m13dc1xdyvr" on amazon.fr,
// and buyer "bm11jdhs75yyts4" on amazon.com.be, whose base local part recurs
// identically across different case tickets (#1000132589 and #1000155203).
//
// Because Zendesk creates a brand-new end-user from the exact From-address
// the FIRST time it sees it, every new case currently spawns a SEPARATE
// end-user whose Primary Email is the full "+uuid" address — so clicking the
// customer's name in Zendesk (which lists tickets by end-user, not by real
// buyer) only ever shows that ONE case, never the buyer's full history across
// orders. Confirmed live: #1000132589's Primary Email is the correct bare
// address; #1000155203 (same real buyer) is a completely different end-user
// whose Primary Email still carries the "+uuid" segment.
//
// Fix: strip the "+uuid" down to the stable base address, then either
//   (a) merge this ticket's just-auto-created end-user into an EXISTING
//       end-user that already has that base address as an identity — so this
//       ticket's history moves onto the one canonical profile, or
//   (b) if no such end-user exists yet, add the base address as a NEW
//       identity on this end-user and make it primary — first time this
//       buyer is seen, this becomes their canonical profile for every future
//       case.
// The non-primary "+uuid" identity is left alone either way — confirmed with
// the team lead that only the Primary Email matters; a secondary address is
// fine to keep or drop.
//
// Mutates `ticket.requester_id` in place when a merge happens (Zendesk
// reassigns the ticket to the merge target server-side) — callers must use
// the ticket object AFTER calling this, not a requester fetched before it.
// Best-effort/non-fatal: any failure here (e.g. a race between two
// near-simultaneous new cases from the same buyer, since Zendesk enforces
// unique identity values account-wide) is caught and logged, never blocks the
// rest of handleNewAbmTicket_'s merge/cleanup/auto-fill for this ticket.
function normalizeAbmRequesterIdentity_(ticket) {
  try {
    const fromAddr = (ticket.via && ticket.via.source && ticket.via.source.from
      && ticket.via.source.from.address) || '';
    const m = fromAddr.match(/^([^@+]+)\+[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}(@marketplace\.amazon\.[\w.-]+)$/i);
    if (!m) return { status: 'no_case_suffix' }; // not an ABM proxy address with a case-id suffix — nothing to normalize

    const baseEmail = (m[1] + m[2]).toLowerCase();
    const requester = getUser_(ticket.requester_id);
    if (!requester) return { status: 'no_requester' };
    if ((requester.email || '').toLowerCase() === baseEmail) {
      return { status: 'already_normalized' }; // a prior run (or Zendesk itself) already fixed this
    }

    const found = zdFetch_(`/api/v2/users/search.json?query=${encodeURIComponent('email:' + baseEmail)}`);
    const canonical = (found.users || []).find(u => u.id !== requester.id);

    if (canonical) {
      zdFetch_(`/api/v2/users/${requester.id}/merge.json`, {
        method: 'put',
        payload: JSON.stringify({ user: { id: canonical.id } })
      });
      ticket.requester_id = canonical.id;
      return { status: 'merged', fromUserId: requester.id, intoUserId: canonical.id, baseEmail };
    }

    const identityRes = zdFetch_(`/api/v2/users/${requester.id}/identities.json`, {
      method: 'post',
      payload: JSON.stringify({ identity: { type: 'email', value: baseEmail, verified: true } })
    });
    const identityId = identityRes.identity && identityRes.identity.id;
    if (identityId) {
      zdFetch_(`/api/v2/users/${requester.id}/identities/${identityId}/make_primary.json`, { method: 'put' });
    }
    return { status: 'primary_identity_added', userId: requester.id, identityId, baseEmail };
  } catch (err) {
    Logger.log(`normalizeAbmRequesterIdentity_: ${err}`);
    return { status: 'error', error: String(err) };
  }
}

// Manual runner/backfill — run from the Apps Script editor. Safe to re-run on
// an already-normalized ticket (returns 'already_normalized', no side effect).
function testNormalizeIdentityOnTicket(ticketId) {
  const ticket = getTicket_(ticketId);
  if (!ticket) { Logger.log('ticket not found'); return { status: 'not_found', ticketId }; }
  const result = normalizeAbmRequesterIdentity_(ticket);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Amazon's own ABM notification email occasionally ships with a broken
// Subject header for certain inquiry types. Confirmed live 2026-08-18: a
// single real JP buyer (藤廣我仁, via.source.from.name correct on every one)
// sent 5 "商品の詳細情報に関するお問い合わせ" messages within about an hour,
// each a genuinely different Seller Central case (different SPC-JPAmazon-…
// id in the body). The FIRST ticket's subject was correct —
// "Amazonのカスタマー 藤廣我仁 様から…に関するお問い合わせ(注文: …)" — every
// later one from the identical sender instead read "…Spigen Japan 様から…"
// (our OWN seller account name, standing in for the buyer's) with the
// "(注文: …)" order suffix dropped entirely. `raw_subject` (Zendesk's copy
// of the unprocessed original header) already showed the broken text, so
// this is Amazon's own template misfiring, not anything in this pipeline —
// but the buyer's real name (from-header) and the order number (always
// present in the body: "# 250-XXXXXXX-XXXXXXX:") are both still reliably
// available elsewhere on the ticket, so there's no reason to leave a
// misleading subject live when it's this cheap to repair.
//
// Originally written narrow (hardcoded 'Spigen Japan' only, JP order-suffix
// phrasing only) — broadened 2026-08-19 after finding a second live instance
// on ticket #1000158113 (Gabor, DE): same bug, but the placeholder was
// "Spigen EU" and the subject's own template was the ENGLISH one ("Product
// details enquiry from Amazon customer …"), so the original hardcoded
// checks silently no-opped on it (`nothing_to_fix`) even though the
// placeholder was unmistakably still one of our own seller account names,
// not a real buyer's. Detecting the placeholder by SHAPE ("Spigen" + at
// most one more word) instead of an exact string covers every marketplace
// variant without needing to enumerate them.
//
// Order-suffix reconstruction is still deliberately narrow — only the JP
// and English templates observed live so far ("…に関するお問い合わせ" +
// "(注文: …)" / "…enquiry from Amazon customer …" + " (Order: …)"). Other
// locales ("Factuuraanvraag van Amazon-klant … (Bestelling: …)", etc.)
// aren't safe to guess the exact phrasing for from here; a subject that
// doesn't match a known template is left untouched rather than risk
// producing a wrong-looking one.
const ABM_SELLER_PLACEHOLDER_RE_ = /^Spigen(\s+\S+)?$/i; // tests a standalone name (e.g. via.source.from.name)
const ABM_SELLER_PLACEHOLDER_SEARCH_RE_ = /Spigen(?:\s+\S+)?/; // finds the same shape sitting inside a subject string
const ABM_SUBJECT_ORDER_TEMPLATES_ = [
  { endsWith: /に関するお問い合わせ$/, hasOrder: /注文[:：]/, build: order => `(注文: ${order})` },
  { endsWith: /\benquiry from Amazon customer\s+\S+.*$/i, hasOrder: /\(Order:/i, build: order => ` (Order: ${order})` },
];
function fixAbmSubjectPlaceholder_(ticket) {
  try {
    const realName = ticket.via && ticket.via.source && ticket.via.source.from && ticket.via.source.from.name;
    if (!realName || ABM_SELLER_PLACEHOLDER_RE_.test(realName)) return { status: 'no_real_name' };

    let subject = ticket.subject || '';
    let changed = false;

    // Find the placeholder token actually sitting in the subject — search
    // for a "Spigen" or "Spigen {word}" run rather than assuming realName's
    // own shape tells us what the BROKEN subject looks like.
    const placeholderMatch = subject.match(ABM_SELLER_PLACEHOLDER_SEARCH_RE_);
    if (placeholderMatch) {
      subject = subject.slice(0, placeholderMatch.index) + realName + subject.slice(placeholderMatch.index + placeholderMatch[0].length);
      changed = true;
    }

    const orderMatch = (ticket.description || '').match(/#\s*(\d{3}-\d{7}-\d{7})/);
    if (orderMatch) {
      const trimmed = subject.trim();
      const template = ABM_SUBJECT_ORDER_TEMPLATES_.find(t => t.endsWith.test(trimmed) && !t.hasOrder.test(trimmed));
      if (template) {
        subject = trimmed + template.build(orderMatch[1]);
        changed = true;
      }
    }

    if (!changed) return { status: 'nothing_to_fix' };

    zdFetch_(`/api/v2/tickets/${ticket.id}.json`, {
      method: 'put',
      payload: JSON.stringify({ ticket: { subject } })
    });
    ticket.subject = subject; // keep in-memory ticket consistent for the rest of this run
    return { status: 'fixed', subject };
  } catch (err) {
    Logger.log(`fixAbmSubjectPlaceholder_: ${err}`);
    return { status: 'error', error: String(err) };
  }
}

// Manual runner/backfill — run from the Apps Script editor.
function testFixAbmSubjectOnTicket(ticketId) {
  const ticket = getTicket_(ticketId);
  if (!ticket) { Logger.log('ticket not found'); return { status: 'not_found', ticketId }; }
  const result = fixAbmSubjectPlaceholder_(ticket);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Case-ID-keyed fast dedup — supplements the search below, which can lag
// several seconds to a minute behind ticket creation (this was a documented
// "known limitation": "a primary created only seconds earlier may not be
// indexed yet ... only matters when two messages arrive nearly
// simultaneously"). Confirmed live 2026-08-10: tickets #1000157138 and
// #1000157139, same Seller Central case id, created 50s apart — the search
// found nothing for the second one (its own primary wasn't indexed yet), so
// both became separate primaries instead of merging. CacheService writes are
// visible to the very next execution almost immediately, unlike Zendesk's
// search index, so this closes that gap without changing the search-based
// path at all when there's no cache hit.
const CASE_PRIMARY_CACHE_PREFIX_ = 'abm_case_primary_';
const CASE_PRIMARY_CACHE_TTL_SEC_ = 3600; // 1hr — comfortably over the observed race window (well under a minute)
function _rememberCasePrimary_(caseId, ticketId) {
  if (!caseId) return;
  try { CacheService.getScriptCache().put(CASE_PRIMARY_CACHE_PREFIX_ + caseId, String(ticketId), CASE_PRIMARY_CACHE_TTL_SEC_); }
  catch (e) { Logger.log('_rememberCasePrimary_: ' + e); }
}
function _lookupCasePrimary_(caseId) {
  if (!caseId) return null;
  try {
    const v = CacheService.getScriptCache().get(CASE_PRIMARY_CACHE_PREFIX_ + caseId);
    return v ? Number(v) : null;
  } catch (e) { Logger.log('_lookupCasePrimary_: ' + e); return null; }
}

// Picks the prior ticket to thread the new message into — see the file header
// for the full selection rationale. Returns the ticket object or null.
function findPrimaryTicket_(requesterEmail, newTicket) {
  const newCaseId = caseIdFromTicket_(newTicket);

  // Fast path: this exact case id already has a known primary in cache —
  // skip the search entirely so this isn't subject to its indexing lag.
  if (newCaseId) {
    const cachedId = _lookupCasePrimary_(newCaseId);
    if (cachedId && cachedId !== newTicket.id) {
      const cached = getTicket_(cachedId);
      if (cached && cached.status !== 'closed') return cached;
    }
  }

  // status<closed keeps new/open/pending/solved (all reusable) and drops only
  // terminal 'closed' tickets, which Zendesk won't let us reopen or comment on.
  const query = encodeURIComponent(
    `type:ticket tags:${ABM_TAG} requester:${requesterEmail} status<closed`
  );
  const data = zdFetch_(`/api/v2/search.json?query=${query}&sort_by=created_at&sort_order=desc`);
  const candidates = (data.results || []).filter(t => t.id !== newTicket.id);
  if (!candidates.length) return null;

  if (newCaseId) {
    const sameCase = candidates.find(t => caseIdFromTicket_(t) === newCaseId);
    if (sameCase) return sameCase;
    // No same-case candidate. Only fall back to candidates with NO resolvable
    // case ID (could plausibly be the same conversation); never merge into a
    // candidate that carries a DIFFERENT explicit case ID.
    const unknownCase = candidates.find(t => caseIdFromTicket_(t) === null);
    return unknownCase || null;
  }
  // New ticket has no case ID → fall back to plain newest same-buyer candidate.
  return candidates[0];
}

function mergeNewTicketIntoPrimary_(newTicket, primaryTicket) {
  const requester = getUser_(newTicket.requester_id);
  const firstComment = getFirstComment_(newTicket.id);
  const htmlBody = (firstComment && firstComment.html_body) || newTicket.description || '';

  // Carry the customer's attachments (photos/video/PDFs) across too — the
  // message body alone isn't enough (reported: 4 photos on #1000153766 didn't
  // reach #1000153709). Re-host them and pass the resulting upload tokens on
  // the primary's comment.
  const uploadTokens = transferAttachments_(firstComment && firstComment.attachments);

  // 1. Add the new message as a public comment on the primary ticket,
  //    authored as the requester so it reads as the customer's follow-up.
  //    Reopen the primary (→ 'open') if it had already been solved, so the
  //    thread comes back to the agents' queue instead of a new ticket.
  const primaryComment = { html_body: htmlBody, public: true, author_id: requester.id };
  if (uploadTokens.length) primaryComment.uploads = uploadTokens;
  const primaryPayload = { comment: primaryComment };
  // Reopen out of ANY waiting state, not just 'solved'. Found live on ticket
  // #1000155020: the primary was 'hold' (agent's ESC/Pending status, waiting
  // on the customer) when a new buyer message merged in — this only checked
  // for 'solved', so 'hold'/'pending' tickets silently stayed put with no
  // queue signal at all; the agent only found out a reply had arrived by
  // separately checking Seller Central. 'new'/'open' are already active in
  // the agent's queue, so nothing to do for those.
  const wasReopened = primaryTicket.status === 'solved' || primaryTicket.status === 'pending' || primaryTicket.status === 'hold';
  if (wasReopened) primaryPayload.status = 'open';

  const mergeResponse = zdFetch_(`/api/v2/tickets/${primaryTicket.id}.json`, {
    method: 'put',
    payload: JSON.stringify({ ticket: primaryPayload })
  });
  // Zendesk's ticket-update response includes an audit trail of what changed;
  // pulling the Comment event's id here is how cleanupExistingAbmTicket_ below
  // is told to clean ONLY this new comment instead of rescanning the whole
  // ticket (see its call site in handleNewAbmTicket_ for why that matters).
  const commentEvent = ((mergeResponse.audit || {}).events || []).find(ev => ev.type === 'Comment');
  const newCommentId = commentEvent ? commentEvent.id : null;

  // 2. Close the duplicate with an internal note pointing to the primary.
  zdFetch_(`/api/v2/tickets/${newTicket.id}.json`, {
    method: 'put',
    payload: JSON.stringify({
      ticket: {
        status: 'closed',
        comment: {
          html_body: `Auto-merged: this buyer already has ABM ticket #${primaryTicket.id}${wasReopened ? ' (reopened)' : ''}. Message moved there — see that ticket for the customer's reply.`,
          public: false
        }
      }
    })
  });

  const srcAttachCount = (firstComment && firstComment.attachments || []).length;
  Logger.log(`Merged ticket #${newTicket.id} into #${primaryTicket.id}${wasReopened ? ' (reopened)' : ''} — attachments ${uploadTokens.length}/${srcAttachCount} transferred`);
  return { wasReopened, attachmentsTransferred: uploadTokens.length, attachmentsTotal: srcAttachCount, newCommentId };
}

// Fast claim guard against Zendesk's own retry-on-timeout behavior. This
// pipeline makes ~10+ sequential Zendesk API calls (identity normalize,
// primary search, merge, cleanup) — enough that a run can legitimately still
// be in progress when Zendesk's webhook response window closes. Zendesk then
// resends the SAME {"ticket_id": N} payload up to 6 times over ~45min. The
// existing `ticket.status === 'closed'` check below already makes a MERGED
// ticket's retries cheap and safe, but a ticket that ends up as its own
// primary (the common case — no prior thread to merge into) never closes, so
// nothing stopped a retry from re-running the ENTIRE expensive chain from
// scratch. Confirmed live via this webhook's invocation history
// (`GET /api/v2/webhooks/{id}/invocations`): total-failure ("terminated")
// count jumped from ~0-1/day the prior week to 12 on 2026-08-10 and 24 on
// 2026-08-11, with no code or ticket-volume change to explain it — consistent
// with retries re-doing the full chain and consuming the very capacity that
// caused the original timeout, a self-reinforcing retry storm. Confirmed two
// concrete casualties: ticket #1000157094 (Federico) never got its Primary
// Email normalized, and #1000157139 never merged into #1000157138 (same
// Seller Central case, 50s apart) — both tickets' webhook invocations show
// every attempt as 'terminated' in the log, never 'success'.
//
// TTL is deliberately SHORT (a few minutes, not the full ~45min retry
// window): most retries in the invocation history DO eventually succeed
// (225/266 recent invocations succeeded, most after 1+ retries) — Zendesk's
// retry mechanism is a real, working recovery path for genuine one-off
// failures, not just noise. A long claim would block that recovery for a
// ticket whose first attempt genuinely died. This TTL only needs to outlast
// one execution's realistic runtime (GAS hard-caps a single execution at a
// few minutes regardless), so it absorbs the retries that land WHILE the
// first attempt is still actually running or has JUST finished, without
// suppressing the later retries that are this ticket's real safety net.
// This claim is deliberately a plain get-then-put (not truly atomic) —
// acceptable here since Zendesk's own retries are spaced 15-45+ seconds
// apart in the observed data, not concurrent.
const ABM_CLAIM_CACHE_PREFIX_ = 'abm_claim_';
const ABM_CLAIM_CACHE_TTL_SEC_ = 300; // 5min

function handleNewAbmTicket_(ticketId) {
  const claimKey = ABM_CLAIM_CACHE_PREFIX_ + ticketId;
  const cache = CacheService.getScriptCache();
  if (cache.get(claimKey)) {
    return { status: 'already_claimed', ticketId };
  }
  cache.put(claimKey, '1', ABM_CLAIM_CACHE_TTL_SEC_);

  const ticket = getTicket_(ticketId);
  if (!ticket) return { status: 'not_found', ticketId };
  if ((ticket.tags || []).indexOf(ABM_TAG) === -1) {
    return { status: 'skipped_no_tag', ticketId };
  }
  // Idempotency guard: this ticket was already processed (closed) by a prior
  // invocation — e.g. a webhook retry. Re-running would re-post its message
  // as a duplicate comment on the primary ticket.
  if (ticket.status === 'closed') {
    return { status: 'already_processed', ticketId };
  }

  // Amazon's own ABM notification email occasionally ships with a broken
  // Subject header for certain inquiry types — see fixAbmSubjectPlaceholder_.
  // Independent of the merge decision below; runs regardless of whether this
  // ticket ends up primary or gets merged away.
  const subjectFix = fixAbmSubjectPlaceholder_(ticket);
  if (subjectFix.status !== 'nothing_to_fix') Logger.log(`fixAbmSubjectPlaceholder_(${ticketId}): ${JSON.stringify(subjectFix)}`);

  // Normalize this buyer's Zendesk end-user identity BEFORE any requester
  // email is read below — see normalizeAbmRequesterIdentity_ for why (Primary
  // Email otherwise carries the case-specific "+uuid" proxy address, which
  // fragments the same real buyer across a separate end-user per case). May
  // mutate ticket.requester_id in place (merge case).
  const identityNormalization = normalizeAbmRequesterIdentity_(ticket);

  const requester = getUser_(ticket.requester_id);
  if (!requester || !requester.email) {
    return { status: 'skipped_no_requester_email', ticketId, identityNormalization };
  }

  const primary = findPrimaryTicket_(requester.email, ticket);
  if (!primary) {
    // This ticket's own first comment is still the raw Amazon template —
    // clean it up here too (see the "ABM inbound cleanup" section above),
    // same as a merged follow-up gets below.
    const cleanup = cleanupExistingAbmTicket_(ticketId);
    // This ticket IS the primary for its case id now — record it immediately
    // so a near-simultaneous duplicate (same case, arriving before Zendesk's
    // search index catches up) finds it via the cache fast-path above instead
    // of also becoming its own primary.
    _rememberCasePrimary_(caseIdFromTicket_(ticket), ticket.id);
    return { status: 'left_as_primary', ticketId, identityNormalization, cleanup };
  }

  // Zendesk's SEARCH index lags behind live ticket state — a search result's
  // `status` can be stale (e.g. still shows 'new' seconds after a solve, or
  // 'open' after a close). The reopen decision and the can-we-write check both
  // depend on the TRUE current status, so re-fetch the chosen primary directly
  // before mutating it. If it's actually already closed (terminal — a PUT
  // would 422 "closed prevents ticket update"), don't merge into it; leave the
  // new ticket as its own primary instead.
  const freshPrimary = getTicket_(primary.id);
  if (!freshPrimary || freshPrimary.status === 'closed') {
    // This ticket ends up as its own primary here too (merge aborted, same
    // as the !primary branch above) — needs the exact same inbound cleanup.
    // Found live 2026-07-30 on ticket #1000155576: this branch returned
    // early with no cleanup call at all, so the ticket permanently kept
    // showing the raw, uncollapsed Amazon-template email with no text-only
    // copy ever generated — a real code gap, not the transient network
    // issue that caused the same symptom on #1000155549 earlier the same
    // day (see the zdFetch_ retry fix above). Confirmed via the webhook's
    // own invocation log: it ran and returned 200 with exactly this status/
    // note, cleanup simply was never invoked.
    const cleanup = cleanupExistingAbmTicket_(ticketId);
    _rememberCasePrimary_(caseIdFromTicket_(ticket), ticket.id);
    return { status: 'left_as_primary', ticketId, note: 'primary was closed on re-fetch', cleanup };
  }

  const merge = mergeNewTicketIntoPrimary_(ticket, freshPrimary);
  // Refresh the cache entry to point at the confirmed live primary — keeps a
  // third near-simultaneous message for this case merging into the right
  // ticket even if this run's own cache write (in the branches above) never
  // happened for it (e.g. this primary predates this fix).
  _rememberCasePrimary_(caseIdFromTicket_(ticket), freshPrimary.id);
  // The comment mergeNewTicketIntoPrimary_ just posted onto the primary
  // carries the SAME raw Amazon template html_body as the source ticket's
  // own first comment — clean it up on the primary right away, same
  // function used for the standalone (left_as_primary) case above. Restricted
  // to JUST that new comment (merge.newCommentId) — without this, a full-
  // ticket scan here would also backfill any OTHER never-cleaned comment
  // already on the primary (e.g. one older than this feature's deploy date),
  // making one new message look like the whole thread arrived again (see
  // ticket #1000153825). Old backlog comments still get cleaned, just via
  // the explicit `cleanupTicket` webhook action / testCleanupOnTicket, not
  // automatically on every new-message trigger.
  const cleanup = cleanupExistingAbmTicket_(freshPrimary.id, merge.newCommentId);
  return {
    status: merge.wasReopened ? 'merged_reopened' : 'merged',
    ticketId,
    primaryTicketId: freshPrimary.id,
    identityNormalization,
    attachmentsTransferred: merge.attachmentsTransferred,
    attachmentsTotal: merge.attachmentsTotal,
    cleanup
  };
}

/********************************
 * Web App entry point (Zendesk Notify-target webhook)
 ********************************/
function doPost(e) {
  try {
    const params = e.parameter || {};
    if (params.secret !== WEBHOOK_SECRET) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'forbidden' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const body = JSON.parse(e.postData.contents || '{}');

    // Secret-guarded on-demand backfill trigger — cleans up a ticket created
    // before the inbound-cleanup feature existed (or any ticket the live
    // webhook path missed for any reason). Idempotent, safe to call repeatedly.
    if (body.action === 'cleanupTicket') {
      const result = cleanupExistingAbmTicket_(body.ticketId);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Secret-guarded on-demand backfill trigger — normalizes a specific
    // ticket's requester's Primary Email to the base (no "+uuid") address,
    // merging into an existing canonical end-user if one already exists. For
    // fixing tickets created before this feature existed.
    if (body.action === 'normalizeIdentity') {
      const ticket = getTicket_(body.ticketId);
      const result = ticket
        ? normalizeAbmRequesterIdentity_(ticket)
        : { status: 'not_found', ticketId: body.ticketId };
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Secret-guarded on-demand backfill trigger — repairs a ticket whose
    // subject Amazon's own ABM email shipped broken (seller-name placeholder
    // in place of the buyer's name, missing order suffix). See
    // fixAbmSubjectPlaceholder_ for the full story. For fixing tickets
    // created before this feature existed.
    if (body.action === 'fixSubject') {
      const ticket = getTicket_(body.ticketId);
      const result = ticket
        ? fixAbmSubjectPlaceholder_(ticket)
        : { status: 'not_found', ticketId: body.ticketId };
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Secret-guarded manual reconciliation trigger (for testing/on-demand runs).
    if (body.action === 'reconcile') {
      // Was `|| 6` — inconsistent with the function's own now-1hr default,
      // so a manual/webhook call with no explicit lookbackHours silently
      // exercised the OLD expensive 6hr window even after that default was
      // reduced (confirmed while testing: this line, not the function's
      // internal fallback, is what a bare {"action":"reconcile"} call with
      // no lookbackHours actually hits, since Number(undefined) || 6 is
      // ALWAYS a number and so never reaches reconcileAbmRelays_'s own
      // typeof-guarded default at all).
      const summary = reconcileAbmRelays_(Number(body.lookbackHours) || 1);
      return ContentService.createTextOutput(JSON.stringify(summary))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Secret-guarded one-time trigger installer.
    if (body.action === 'setupTrigger') {
      const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'reconcileAbmRelays_');
      if (!exists) ScriptApp.newTrigger('reconcileAbmRelays_').timeBased().everyMinutes(30).create();
      return ContentService.createTextOutput(JSON.stringify({ installed: !exists, alreadyExisted: exists }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ticketId = body.ticket_id;
    if (!ticketId) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'missing ticket_id' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result = handleNewAbmTicket_(ticketId);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    // Zendesk's Notify-target webhook only sees this function's HTTP 200
    // response — it has no way to know handleNewAbmTicket_ actually threw,
    // so it never retries and the ticket is silently orphaned (never
    // merged/reopened into its primary, subject never corrected). Confirmed
    // live 2026-08-18/19: a same-day Google Workspace UrlFetchApp daily
    // quota exhaustion ("premium urlfetch 서비스를 너무 많이 호출") caused
    // exactly this for at least 2 real customers (Gabor #1000158113, Mandeep
    // #1000158120) — agents only found out by noticing the broken subject/
    // missing merge and fixing it by hand. Logger.log alone isn't visible to
    // anyone without digging through the Apps Script Executions dashboard —
    // tag the actual ticket instead so it's findable via a saved view and
    // re-processable once the underlying issue (e.g. quota reset) clears.
    // Best-effort: if THIS also fails (plausible for the same reason), we're
    // no worse off than before — never let it mask the original error.
    try {
      const failedTicketId = (JSON.parse(e.postData.contents || '{}')).ticket_id;
      if (failedTicketId) {
        // POST .../tags.json is additive (unlike PUT .../tickets/{id}.json,
        // which would REPLACE the ticket's whole tag set) — safe to fire
        // without first reading back the ticket's current tags.
        zdFetch_(`/api/v2/tickets/${failedTicketId}/tags.json`, {
          method: 'post',
          payload: JSON.stringify({ tags: ['abm_auto_process_failed'] })
        });
        zdFetch_(`/api/v2/tickets/${failedTicketId}.json`, {
          method: 'put',
          payload: JSON.stringify({
            ticket: { comment: { public: false, body: `자동 병합/정리 처리 실패 (재시도 필요): ${String(err)}` } }
          })
        });
      }
    } catch (tagErr) {
      Logger.log(`doPost failure-tagging also failed: ${tagErr}`);
    }
    return ContentService.createTextOutput(JSON.stringify({ error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/********************************
 * Manual test helper — run from the Apps Script editor
 * (bypasses the webhook/secret entirely; calls the same logic directly)
 ********************************/
function testMergeOnTicket(ticketId) {
  const result = handleNewAbmTicket_(ticketId);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/********************************
 * ABM relay reconciliation
 * ------------------------------
 * Safety net for the outgoing relay. GCX Reply relays an agent's public reply
 * to the buyer via Seller Central ONLY from a browser running the updated
 * userscript with a live SC session. If an agent's browser is stale/old/not
 * logged into SC, their reply silently never reaches the buyer AND never gets
 * logged (the relay never ran to log a failure) — exactly what happened on
 * ticket #1000153794 (agent sent an invoice reply + 3 PDFs; buyer never got
 * it; zero rows in ABM_Relay_Log).
 *
 * This scheduled job closes that gap: it scans recently-updated ABM tickets,
 * finds each ticket's latest public AGENT reply, and if that reply isn't
 * already accounted for in ABM_Relay_Log (by CommentId, or by matching text
 * for browser-relayed rows that predate CommentId), it queues a 'pending' row
 * in the log. Any updated agent browser with a valid SC session then claims
 * (atomic) and delivers it — text AND attachments — via the existing sweep.
 *
 * Dedup is deliberately conservative (favours "already handled" to avoid a
 * duplicate send to a real buyer). Sends are NOT performed here — this only
 * queues; the browser sweep does the actual delivery, one claimant per row.
 ********************************/

// GCX Reply web app (owns the ABM_Relay_Log sheet + its endpoints).
const GCX_GAS_URL = 'https://script.google.com/macros/s/AKfycbw2Vdwk197LXB6oUAzuHS8sKamD5uqKZJDLvcHzbftWJk-M65XV1fAnTqiZo7ZEm4hk/exec';

// Minimum age (from the reply's own created_at) before reconciliation will
// consider it — see the usage site in reconcileAbmRelays_ for the race this
// closes (reconciliation running before the live browser relay has finished
// logging its own attempt). Comfortably longer than the live relay's own
// retry budget (up to 3 attempts with backoff) could plausibly take.
const ABM_RECONCILE_MIN_AGE_MS = 10 * 60 * 1000; // 10 minutes

// EU marketplaces whose Seller Central messaging is served from amazon.de
// (single EU login covers them via marketplaceId) — mirrors GCX Reply's
// EU_SC_REDIRECT so the queued row's marketplace matches what the browser
// sweep will actually hit.
const EU_SC_REDIRECT_TO_DE = ['amazon.fr', 'amazon.it', 'amazon.es', 'amazon.nl', 'amazon.pl', 'amazon.se', 'amazon.be'];

function abmMarketplaceFromAddress_(address) {
  const m = (address || '').match(/@marketplace\.(amazon\.[a-z.]+)$/i);
  if (!m) return null;
  const domain = m[1].toLowerCase();
  return EU_SC_REDIRECT_TO_DE.indexOf(domain) !== -1 ? 'amazon.de' : domain;
}

// Aggressive normalization for dedup: the browser relay stores the DECODED
// message text (htmlToPlainText_ turns &nbsp; into an actual nbsp char, etc.),
// while Zendesk's comment.plain_body keeps literal HTML entities like
// "&nbsp;". Comparing raw would falsely treat the same reply as different and
// risk a duplicate send. So strip HTML entities first, then drop everything
// that isn't a letter or digit, then lowercase — two encodings of the same
// message collapse to the same key. Erring toward "already handled" (a false
// match just skips a resend) is the safe direction for real-customer sends.
function normalizeForDedup_(s) {
  return String(s || '')
    .replace(/&[a-z]+;|&#\d+;/gi, ' ')   // HTML entities → space (so "nbsp" letters don't linger)
    .replace(/[^\p{L}\p{N}]+/gu, '')     // keep only letters/digits
    .toLowerCase();
}

// Retries on a non-JSON (HTML) response — confirmed live 2026-08-19 that
// GCXReply_GAS (a separate, independently heavily-loaded project — every
// open agent browser tab hits it every 5min, plus live panel usage)
// intermittently returns Google's own platform error page instead of its
// real JSON output, almost certainly its own concurrent-execution limit,
// not anything wrong with the request itself. A concurrency LOCK on the
// ABM_TicketMerge caller side (tried first) did NOT fix this — confirmed
// by watching a run fail with the identical error while running the
// locked code, proving the overlap (if any) isn't on this side. Since
// Zendesk's own retry mechanism for analogous transient failures succeeds
// the vast majority of the time (see reconcileFailedAbmProcessing_'s doc
// comment / [[abm_ticket_merge]] v29/30), a short retry here is the
// correctly-targeted fix — give GCXReply_GAS's concurrent slot a moment
// to free up rather than crashing the whole reconcile run on the first
// hiccup.
const GCX_FETCH_RETRIES = 3;
function gcxFetch_(url, options) {
  for (let attempt = 1; attempt <= GCX_FETCH_RETRIES; attempt++) {
    const res = UrlFetchApp.fetch(url, Object.assign({ muteHttpExceptions: true, followRedirects: true }, options || {}));
    const body = res.getContentText();
    try {
      return JSON.parse(body);
    } catch (e) {
      if (attempt < GCX_FETCH_RETRIES) { Utilities.sleep(1000 * attempt); continue; }
      throw new Error(`gcxFetch_ ${url} -> non-JSON response after ${GCX_FETCH_RETRIES} attempts: ${body.slice(0, 200)}`);
    }
  }
}

function gcxGet_(action, params) {
  let url = `${GCX_GAS_URL}?action=${encodeURIComponent(action)}`;
  Object.keys(params || {}).forEach(k => { url += `&${k}=${encodeURIComponent(params[k])}`; });
  return gcxFetch_(url);
}

function gcxPost_(payload) {
  return gcxFetch_(GCX_GAS_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) });
}

// Latest PUBLIC agent reply on a ticket — i.e. an outbound message that should
// have reached the buyer. Excludes the customer's own inbound messages and the
// merge-injected customer follow-ups (both authored by end-users, not agents).
function latestAgentPublicComment_(ticketId, requesterId) {
  const data = zdFetch_(`/api/v2/tickets/${ticketId}/comments.json?include=users&sort_order=desc`);
  const roleById = {};
  (data.users || []).forEach(u => { roleById[u.id] = u.role; });
  const comments = data.comments || [];
  for (const c of comments) {
    if (!c.public) continue;
    if (c.author_id === requesterId) continue;             // customer / merge-injected
    const role = roleById[c.author_id];
    if (role === 'agent' || role === 'admin') return c;    // outbound agent reply
  }
  return null;
}

function reconcileAbmRelays_(lookbackHours) {
  // Time-based triggers invoke this with an event object as the first arg, not
  // a number — only honour a real numeric override (from runReconcileNow /
  // the `reconcile` webhook action's explicit lookbackHours param).
  //
  // Default was 6 hours — meaning the UNATTENDED 30-min CLOCK trigger (the
  // only caller that ever hits this fallback) re-scanned a full 6-HOUR
  // window of ABM tickets on EVERY run, up to 12x redundant overlap before
  // a ticket aged out of the window. Confirmed live 2026-08-19 via the Apps
  // Script executions dashboard: with the real 6hr default, a single run
  // took 265s (measured directly); the dashboard showed repeated 170-362s
  // runs and outright failures clustered minutes apart — right at GAS's
  // 6-minute (360s) execution ceiling, some pushed over it. This redundant
  // rescanning (same tickets re-checked ~12x before aging out) is a
  // plausible major contributor to the account-wide UrlFetchApp daily quota
  // exhaustion this session's other fixes were built around (see
  // [[abm_ticket_merge]] v31-v34) — not just a latency problem.
  //
  // ABM_RECONCILE_MIN_AGE_MS (10min) already gates against acting on a
  // reply too young for the LIVE relay to have had its own chance first;
  // 1hr gives ~2x overlap over the 30-min trigger cadence (comfortable
  // safety margin against timing drift/a single missed run) instead of 12x.
  // This function is a SAFETY NET for the independent live relay path
  // (relayAbmReply_, fires instantly on an agent's reply) — a narrower
  // backup window is an acceptable tradeoff against the redundant-scan cost
  // it was actually paying for.
  // Concurrency guard: confirmed live 2026-08-19 that this trigger is
  // firing far more often than its intended 30-min cadence (cause still
  // unresolved — only one trigger is registered, ruling out duplicates),
  // and each run takes 30-150s+. That means multiple invocations were
  // genuinely OVERLAPPING in wall-clock time, all hammering GCXReply_GAS
  // concurrently via gcxGet_/gcxPost_ below — confirmed as the direct
  // cause of repeated failures in the executions dashboard: `SyntaxError:
  // Unexpected token '<', "<!DOCTYPE "... is not valid JSON at gcxGet_`,
  // i.e. GCXReply_GAS returning Google's own platform error page (almost
  // certainly its concurrent-execution limit) instead of a real response.
  // A non-blocking script lock means an overlapping second invocation
  // bails out immediately instead of piling on more concurrent load — the
  // next firing (whenever that is) picks up the same recent-tickets
  // window anyway, so skipping one redundant overlapping run costs
  // nothing real.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('reconcileAbmRelays_: skipped, another invocation is already running');
    return { skipped: 'already_running' };
  }
  try {
    return reconcileAbmRelays__locked_(lookbackHours);
  } finally {
    lock.releaseLock();
  }
}

function reconcileAbmRelays__locked_(lookbackHours) {
  const hours = (typeof lookbackHours === 'number' && lookbackHours > 0) ? lookbackHours : 1;
  const cutoff = new Date(Date.now() - hours * 3600 * 1000).toISOString().replace(/\.\d+Z$/, 'Z');
  const query = encodeURIComponent(`type:ticket tags:${ABM_TAG} updated>${cutoff}`);

  const summary = { scanned: 0, queued: 0, skippedNoCase: 0, skippedNoReply: 0, skippedTooRecent: 0, alreadyLogged: 0, queuedTickets: [] };

  let url = `/api/v2/search.json?query=${query}&sort_by=updated_at&sort_order=desc`;
  let guard = 0;
  while (url && guard++ < 10) {
    const page = zdFetch_(url);
    (page.results || []).forEach(t => {
      if (!t || t.id === undefined) return;
      summary.scanned++;
      const fromAddr = (t.via && t.via.source && t.via.source.from && t.via.source.from.address) || '';
      const caseId = caseIdFromTicket_(t);
      const marketplace = abmMarketplaceFromAddress_(fromAddr);
      if (!caseId || !marketplace) { summary.skippedNoCase++; return; }

      const reply = latestAgentPublicComment_(t.id, t.requester_id);
      if (!reply) { summary.skippedNoReply++; return; }

      // A SEPARATE bug from the dedup-window one above: the LIVE browser
      // relay (relayAbmReply_) fires the instant an agent's reply posts, but
      // its round-trip (send + log) takes real time — anywhere from a couple
      // seconds to longer with retries. If this reconciliation run's own
      // scheduled tick happens to land in that gap, its dedup check (however
      // thorough) finds NOTHING yet, because the live relay simply hasn't
      // logged its own attempt yet — so it queues a genuine duplicate that
      // some browser then actually sends. Confirmed live via ABM_Relay_Log
      // (2026-07-30): 22 real tickets got double-sent this way, with gaps as
      // small as 6 seconds between the live relay's send and reconciliation's
      // duplicate (e.g. #1000154373, #1000154472, #1000154485) — no amount of
      // improving the DEDUP CHECK closes this, since the race is about
      // WHETHER a row exists yet at all, not whether it's visible once it
      // does. Fix: never consider a reply for reconciliation until it's had a
      // safe head start — comfortably longer than the live relay's own retry
      // budget (up to 3 attempts with backoff) could plausibly take. A reply
      // that's still this young will simply be reconsidered on next run.
      const replyAgeMs = Date.now() - new Date(reply.created_at).getTime();
      if (replyAgeMs < ABM_RECONCILE_MIN_AGE_MS) { summary.skippedTooRecent++; return; }

      const text = (reply.plain_body || '').trim();
      if (!text) { summary.skippedNoReply++; return; }

      // Per-ticket, UNTRUNCATED history — not a slice of the account-wide
      // log. Previously this dedup checked against only the 200 most-recent
      // rows account-wide (one bulk `abmRelayAll` fetch before this loop),
      // sorted by timestamp globally. With ABM_Relay_Log at 470+ rows and
      // growing, any ticket whose relay happened more than ~200 relay-events
      // ago (days, not months, at current volume) had its row silently
      // pushed out of that window — so if the ticket's `updated_at` got
      // bumped again for ANY reason within this run's lookback window (e.g.
      // Zendesk's own automatic close-after-N-days-solved rule, which counts
      // as an update), the dedup check saw an empty `existing` list and
      // re-queued an ALREADY-DELIVERED reply as if it had never been sent —
      // a genuine duplicate message to a real customer. Confirmed live and
      // reported by an agent on ticket #1000154135 (2026-07-29): the original
      // relay row from 2026-07-19 ranked #383 by recency among 473 total log
      // rows, entirely outside the old top-200 window; the ticket auto-closed
      // today, reconciliation re-scanned it, found no matching row, and
      // re-sent the same 10-day-old reply. (The text-normalization dedup
      // itself was verified correct — &nbsp; vs literal whitespace and
      // &amp; vs literal & both normalize identically; this was purely a
      // window-size/scale bug, not a normalization bug.)
      const existing = (gcxGet_('abmRelayStatus', { ticketId: t.id }).status) || [];
      const norm = normalizeForDedup_(text);
      const dup = existing.some(r =>
        String(r.commentId) === String(reply.id) ||          // same comment already logged
        normalizeForDedup_(r.messageText) === norm            // browser-relayed row (no commentId) with same text
      );
      if (dup) { summary.alreadyLogged++; return; }

      gcxPost_({
        action: 'logAbmRelay',
        relayKey: `${t.id}_c${reply.id}`,
        ticketId: t.id,
        commentId: reply.id,
        caseId: caseId,
        marketplace: marketplace,
        status: 'pending',
        attempts: 0,
        lastError: 'queued by reconciliation (relay never logged a result)',
        messageText: text
      });
      summary.queued++;
      summary.queuedTickets.push(t.id);
    });
    url = page.next_page || null;
  }

  Logger.log('reconcileAbmRelays_: ' + JSON.stringify(summary));

  // Piggyback on this same 30-min trigger rather than installing a second
  // one — recovers tickets whose auto-processing threw (doPost's catch
  // block tags them abm_auto_process_failed; see that comment for the
  // 2026-08-18/19 UrlFetchApp daily-quota incident this exists for) once
  // whatever blocked them clears. Best-effort: a failure here must never
  // prevent the ABM relay reconciliation above from having already run.
  try {
    const failedProcessing = reconcileFailedAbmProcessing_();
    summary.failedProcessing = failedProcessing;
  } catch (e) {
    Logger.log('reconcileFailedAbmProcessing_ call from reconcileAbmRelays_ failed: ' + e);
  }

  // Same piggyback pattern, same reasoning: best-effort, never blocks the
  // relay reconciliation above. See reconcileAbmCleanups_ for what this covers.
  try {
    const cleanups = reconcileAbmCleanups_();
    summary.cleanups = cleanups;
  } catch (e) {
    Logger.log('reconcileAbmCleanups_ call from reconcileAbmRelays_ failed: ' + e);
  }

  return summary;
}

// Safety net for the INBOUND TEXT CLEANUP path (cleanupExistingAbmTicket_ /
// handleNewAbmTicket_) — a different gap from reconcileAbmRelays_ above,
// which only covers the Seller-Central-relay log. cleanupExistingAbmTicket_
// guards its work with a single SCRIPT-WIDE LockService lock
// (`lock.tryLock(20000)`) — when a buyer sends multiple ABM messages close
// together (confirmed live down to 1 second apart), their separate
// webhook-triggered cleanup calls compete for that one lock, and a losing
// invocation just returns `{status:'locked'}` with NO retry — that
// comment's clean-text copy is then never generated, permanently (GCX
// Reply correctly leaves it as the full raw "Amazon card" rather than
// hiding it, per the v3.6.18 fix, but the customer's message stays
// effectively unreadable without clicking through). Confirmed live
// 2026-09-21 on ticket #1000162353: 4 raw comments with zero clean pair,
// each immediately followed (1s–91s later) by ANOTHER raw comment that got
// its own pair fine seconds later — the lock-race signature; one stuck
// since 2026-09-18, still uncleaned 3 days later with no automatic
// recovery, since nothing previously re-checked already-seen tickets for
// this specific gap.
//
// Reuses cleanupExistingAbmTicket_ itself rather than reimplementing its
// raw-detection/dedup logic — that function already no-ops instantly (via
// alreadyCleanedSourceIds_) for any comment already cleaned, so calling it
// on every recently-active ABM ticket every 30 min is safe and cheap for
// the common case where nothing is actually missing. If this sweep's own
// LockService call loses the race too (e.g. running concurrently with a
// live webhook), it just returns 'locked' and self-heals on the next tick,
// same tradeoff reconcileAbmRelays_ already accepts.
function reconcileAbmCleanups_(lookbackHours) {
  const hours = (typeof lookbackHours === 'number' && lookbackHours > 0) ? lookbackHours : 1;
  const cutoff = new Date(Date.now() - hours * 3600 * 1000).toISOString().replace(/\.\d+Z$/, 'Z');
  const query = encodeURIComponent(`type:ticket tags:${ABM_TAG} updated>${cutoff}`);

  const summary = { scanned: 0, cleaned: 0, locked: 0, errors: 0, cleanedTickets: [] };
  let url = `/api/v2/search.json?query=${query}&sort_by=updated_at&sort_order=desc`;
  let guard = 0;
  while (url && guard++ < 10) {
    const page = zdFetch_(url);
    (page.results || []).forEach(t => {
      if (!t || t.id === undefined) return;
      summary.scanned++;
      try {
        const result = cleanupExistingAbmTicket_(t.id);
        if (result && result.status === 'locked') { summary.locked++; return; }
        if (result && result.cleaned && result.cleaned.length) {
          summary.cleaned += result.cleaned.length;
          summary.cleanedTickets.push(t.id);
        }
      } catch (e) {
        summary.errors++;
        Logger.log(`reconcileAbmCleanups_(${t.id}): ${e}`);
      }
    });
    url = page.next_page || null;
  }
  Logger.log('reconcileAbmCleanups_: ' + JSON.stringify(summary));
  return summary;
}

// Recovers tickets tagged abm_auto_process_failed (see doPost's catch
// block) by simply retrying handleNewAbmTicket_ once whatever blocked the
// original attempt has cleared — e.g. the account-wide UrlFetchApp daily
// quota (resets every 24h) that caused this for real customers Gabor
// (#1000158113) and Mandeep (#1000158120) on 2026-08-18, silently orphaned
// with no automatic recovery until this existed. handleNewAbmTicket_'s own
// 5-min claim-guard cache means calling it again here is always a REAL
// retry by the time this runs (30-min cadence, well past that window), not
// a race with the original attempt. On success, removes the tag (Zendesk's
// tags.json DELETE is additive-safe — only removes the named tag, doesn't
// touch any others) and leaves a note; on repeat failure, leaves the tag
// for the next sweep and just logs — no retry cap, since a stuck ticket
// retried every 30 min indefinitely is cheap and self-evidently visible
// via the tag either way.
function reconcileFailedAbmProcessing_() {
  const summary = { scanned: 0, recovered: 0, stillFailing: 0, recoveredTickets: [], stillFailingTickets: [] };
  const query = encodeURIComponent('type:ticket tags:abm_auto_process_failed');
  let url = `/api/v2/search.json?query=${query}&sort_by=created_at&sort_order=asc`;
  let guard = 0;
  while (url && guard++ < 10) {
    const page = zdFetch_(url);
    (page.results || []).forEach(t => {
      if (!t || t.id === undefined) return;
      summary.scanned++;
      try {
        handleNewAbmTicket_(t.id);
        zdFetch_(`/api/v2/tickets/${t.id}/tags.json`, {
          method: 'delete',
          payload: JSON.stringify({ tags: ['abm_auto_process_failed'] })
        });
        zdFetch_(`/api/v2/tickets/${t.id}.json`, {
          method: 'put',
          payload: JSON.stringify({ ticket: { comment: { public: false, body: '자동 병합/정리 재처리 성공 (reconcileFailedAbmProcessing_)' } } })
        });
        summary.recovered++;
        summary.recoveredTickets.push(t.id);
      } catch (err) {
        Logger.log(`reconcileFailedAbmProcessing_: ticket ${t.id} still failing: ${err}`);
        summary.stillFailing++;
        summary.stillFailingTickets.push(t.id);
      }
    });
    url = page.next_page || null;
  }
  if (summary.scanned) Logger.log('reconcileFailedAbmProcessing_: ' + JSON.stringify(summary));
  return summary;
}

// Manual runner (dry-run visibility) — run from the Apps Script editor.
function runReconcileNow() {
  return reconcileAbmRelays_(24);
}

// Install the scheduled reconciliation (run ONCE from the editor).
function setupReconcileTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'reconcileAbmRelays_');
  if (exists) { Logger.log('reconcile trigger already exists'); return; }
  ScriptApp.newTrigger('reconcileAbmRelays_').timeBased().everyMinutes(30).create();
  Logger.log('reconcile trigger created — every 30 min');
}

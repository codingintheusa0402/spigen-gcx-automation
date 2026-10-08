# ticket-reporter — full current logic (as of 2026-09-22, post all today's /revision fixes)

> Compiled for review. Source of truth is `SKILL.md` (927 lines) — this is a condensed
> map of it, not a replacement. Flagged items (⚠️) are things worth a decision.

---

## ⚠️ TOP ISSUE FOUND WHILE COMPILING THIS

**Every report sent this session used the wrong greeting line.**

- **Documented rule** (SKILL.md, dated 2026-09-16 — predates this session):
  `리더님 & 프로님들, 아래 티켓 '<조치>' 처리 제안드리며 열람 부탁드립니다. 감사합니다!`
  (explicitly says the old line below was "폐기" — discarded)
- **What I actually sent, every time today**:
  `프로님, 아래 티켓 '...' 처리 컨펌 부탁드립니다. 감사합니다!`

Root cause: I was pattern-matching off earlier turns in this same conversation instead of
re-reading SKILL.md's own template each time — the wrong phrasing got "locked in" early
and repeated ~13 times (#1000162927, #1000162859, #1000162836, #1000162934, #1000162648,
#1000162079, #1000162942, #1000163010, #1000162483, #1000162847, #1000162353 ×2, plus the
Lazada TCT one). Needs a decision: fix going forward only, or also send corrections for
the already-sent ones?

---

## 1. Trigger & cadence

- **Zendesk views**: every 5 min (cron `*/5 * * * *`, job `c572e4be`, 7-day auto-expiry).
- **TCT log** (Lazada/Shopee): every 30 min, throttled via `state/processed.json`'s
  `tct_log_last_scan` unix timestamp — checked inside the *same* 5-min tick, not a
  separate schedule.
- **Business hours gate (first check, every tick)**: weekday KST 08:30–18:00 only. Outside
  that window the tick does *nothing* — no `/revision` check, no Zendesk scan, no TCT scan.
- **`/revision` feedback check (first real step inside business hours)**:
  `check_feedback.py list` → apply each unapplied row to SKILL.md → `mark-applied <row>`.
  ⚠️ **This step was not actually run at the start of most ticks today** — I only
  discovered the two pending rows when you asked directly. This should run automatically
  every tick, not just when prompted.

## 2. Zendesk source — gating

For each row in both views:
1. **Category column must read exactly `4. Product Issue`** (regex `/^4\.\s*Product Issue$/i`).
   Never inferred from tags/title — view column only.
2. **Ticket status column must read exactly `Pending`.** On-hold/Open/New/Solved never
   qualify even if Category matches.
3. **ESC. 사유 exclusions** (skip without opening the ticket, log to `skipped_*`, don't
   add to `processed` so it's re-checked if ESC. 사유 later changes):
   - `MCF발송` → `skipped_mcf`
   - `출시/재입고 문의` → `skipped_launch_inquiry`
   - `유관부서 이관 요청` → `skipped_dept_transfer`
4. **Dedup / reentry logic** via `zendesk_last_status` map (previous tick's observed
   status per ticket ID):
   - Not in `processed` → new, generate + send.
   - In `processed`, currently Pending, previous status **was not** Pending (excursion
     to On-hold/Open/Solved and back) → reentry, generate + send again.
   - In `processed`, currently Pending, previous status **was already** Pending → skip,
     log `skip_dup_reentry`.
   - No prior `zendesk_last_status` entry at all → treat as "was not Pending", allow one
     resend (self-heals the tracking from then on).
   - `zendesk_last_status` is updated to the current observation *after* the above
     decision (order matters).

## 3. TCT log source (Lazada/Shopee) — separate gating

- Tab structure: `Lazada log` / `Shopee log`, header row 4, data from row 5, columns A–R.
- **Target row**: column A (Status), trimmed, exactly `Esc T2`. No Category gate — being
  in `Esc T2` at all is itself the Tier-2-needs-review signal.
- **Dedup**: `processed_tct_log` (full history) + `tct_log_state` (per-key
  `cycled_through_t1` flag):
  - New key → send, register `cycled_through_t1: false`.
  - Known key, current status `Esc T1` → advice was given; flip `cycled_through_t1: true`,
    don't resend.
  - Known key, current status `Esc T2` again, `cycled_through_t1` was `true` → genuine
    re-escalation, resend, reset flag to `false`.
  - Known key, current status `Esc T2`, flag still `false` → already sent, still waiting
    on a human reply; skip.
- **Attachments**: R-column cell's `hyperlink` (not display text) → branch on
  `/folders/` vs `/file/d/` → for images use `thumbnailLink` resized to `s1600` (NOT
  `webContentLink`, which doesn't render in Chat's image widget due to a 303 redirect the
  widget won't follow).

## 4. Per-ticket research (before writing anything)

1. Open the ticket, read the **entire** conversation — customer messages, public replies,
   internal notes — to the actual last item on the page.
   ⚠️ **HARD RULE added today**: never conclude a `예상 답변:` note doesn't exist without
   scrolling all the way down first. Caught twice today (#1000162079, #1000162942) —
   in both cases an internal note with a real `예상 답변:` line existed further down than
   where I'd stopped reading.
2. Extract: country, Device/Product Name (→ ASIN/SKU), purchase date, per-turn
   claim/response history, 최근 2년 order/refund stats, and the **last** `예상 답변:` note
   verbatim (see §5).
3. Collect attachments — only genuinely issue-relevant customer/TCK images (not avatars,
   logos, `~WRD0000.jpg`-style paste artifacts). Includes pasted (ctrl/cmd+V, no download
   link) images — these need a `document.querySelectorAll('img')` src scrape since the
   accessibility tree alone won't show their URL.
   ⚠️ **HARD RULE added today**: on a reentry, still re-collect and attach the customer's
   existing relevant photos from *earlier* turns even if the newest turn added nothing new
   — "not new" is not a reason to omit evidence that's still backing the case.

## 5. Format decision — which report shape to use

**Default: standard 5-part format** (`[Ticket Info.]`, `[문의 요약]`, `[비고]`,
`[TCK 예상 답변]`, `[GCX AI 예상답변]`).

- **`[TCK 예상 답변]`**: copied **verbatim** from the ticket's own internal note — never
  paraphrased, never reordered, never invented. If multiple `예상 답변:` notes exist, use
  the **last** one, and re-verify nothing comes after it before locking that in. If truly
  none exists anywhere on the fully-scrolled page → `확인 불가` (not a guess).
- **`[GCX AI 예상답변]`**: 3-line structure (유감/처리한계/조치) or `동일 (근거)` if it
  matches TCK's answer — `동일` alone is never enough, always needs a parenthetical reason.

**`[처리 답변]` variant** (replaces the `[TCK 예상 답변]`+`[GCX AI 예상답변]` pair with one
"already decided" section): ⚠️ **HARD RULE tightened today** — use this **only** when the
user explicitly states the case is already fully decided (as in the original one-off
#1000162353 request). Never infer this myself from a public reply already having gone out
to the customer — a `예상 답변:` internal note existing means TCK is still asking GCX to
confirm, regardless of whether a public reply already shipped. Real example today
(#1000162079): TCK's note proposed "100% 환불" but the actual GCX decision already written
into the ticket was the opposite ("도움 불가") — a public reply alone tells you nothing
about whether GCX has actually weighed in.

**TCT/Lazada-Shopee variant** (different field mapping, see §3's source and SKILL.md's
"리포트 작성 — 필드 매핑" section): no ASIN, SKU is hyperlinked to a product-name search
(not the SKU itself — SKU-only search returns unrelated products), 구매일자 from column E
(`N/A` if blank, not `확인 불가`), header renamed `[TCT 예상 답변]` (almost always
"해당없음" — this source has TCT asking GCX, not proposing), `[GCX AI 예상답변]` is 2 lines
in **English** (voucher tier + one-sentence reason) since the reader is a non-Korean-
speaking TCT agent, and the CS policy is more lenient (default to 100%/50% voucher even
without confirmed warranty coverage; 10%/decline only for clear abuse signals).

## 6. Content rules (apply regardless of format)

- Never admit/confirm a manufacturing defect, even hedged. Only ever "확인 불가" /
  "사료됨" framing — banned phrases enumerated in SKILL.md §표현주의사항 1.2.
- Loyalty (`충성 고객`): only stated when total orders ≥3 **and** refunds ≤ 1/3 of orders.
  Below 3 orders → never loyal regardless of refund count, and the whole line is omitted
  (not written as "충성 고객 아님") — loyalty is a positive-only callout.
- Abuse signal: refund ratio ≥50% with ≥3 orders → called out explicitly.
- `Legal Action 언급` (law/consumer-rights citations, threats to escalate): classified as
  강성고객 (difficult customer) signal only — never weighed as real legal risk in the
  recommended action.
- Same-customer history cited as `동일고객 클레임 이력:` (not `유사 전례:`, which is
  reserved for a *different* customer's precedent).
- NRN: "도움 불가 N차 안내" caps at N=3. Beyond that, plain `NRN` — never "4차" etc.
- FBA counterfeit/used-item claims → always conclude FBA CS 교환/환불 안내, never
  예외적 환불, regardless of what TCK's note suggests.
- 황변 (yellowing) claims → must include TCK-list item 14's phrasing.
- Purchase-date highlighting: two independent thresholds — `[비고]`'s "+구매 NN개월
  경과" line turns red past **2 months**; `[Ticket Info.]`'s `구매일자:` date itself turns
  red past **3 months**. Both apply independently; don't conflate them.
- Keyword coloring: red (`#D93025`) for decline/no-help/warranty-excluded/NRN/abuse/legal-
  mention keywords and the actual defect-symptom noun (from a ~110-term list); blue
  (`#1A73E8`) for refund/replacement/MCF/exception-granted keywords. Hedge phrases like
  "제조상 결함 확인 불가" are never colored — only concrete symptom/action keywords are.

## 7. Rendering rules

- cardsV2 only, no separate text line above the card.
- Line spacing: `<br>` alone between fields that belong together (Ticket Info's 4 lines,
  비고's two `+` lines, GCX AI 예상답변's 1/2/3 lines); `<br><br>` between sections and
  between list items within a section. Never 2+ blank lines.
- All `[Header]` lines bolded. All ticket-number mentions and the `링크:` URL are
  hyperlinked to the Zendesk ticket.
- HTML-escape `<`, `>`, `&`.

## 8. Sending

- `send.py --html body.html --ticket-id <id> [--image url]... [--attach "name|url"]...`
- `--ticket-id` is **mandatory** on every real send — omitting it silently breaks the
  thread-reply mapping (`TicketQueue` sheet) with no visible error.
- Monitor mode sends directly, no `--dry-run` confirmation step (explicit standing user
  instruction). Manual/single-ticket requests dry-run first.
- PII (bank account numbers, etc.) mentioned in a ticket is never copied into the Chat
  report even if visible on the ticket page — real example today: #1000163010's IBAN.

## 9. Post-send bookkeeping

- `state/processed.json`: `processed` (sent ticket IDs), `zendesk_last_status`,
  `skipped_non_pi` / `skipped_mcf` / `skipped_launch_inquiry` / `skipped_dept_transfer`,
  `log` (per-send audit trail incl. `skip_dup_reentry` / `outside_business_hours` events),
  `processed_tct_log`, `tct_log_state`, `tct_log_last_scan`.
- Zendesk ticket status is **never** touched by the monitor (HARD RULE since 2026-09-17) —
  only a human's Chat thread reply moves it, via the separate Chat-app flow.
- **SIREN check** runs after every successful send (see §10) regardless of outcome.

## 10. SIREN check (defect-pattern early warning)

Runs once per sent ticket, right after send, never retried on failure (just logs a skip).

1. **Allowlist gate**: defect term must be in `siren.py`'s ~75-term `ALLOWED_DEFECT_TERMS`
   (manufacturing-defect-shaped terms only — shipping/mishap terms like 황변,
   동봉제품상이함, 오배송, 기타사항 are explicitly excluded).
2. **신제품 라인업 gate**: the ticket's Device+Product must have a matching row in the
   `신제품 라인업` sheet (with a fuzzy fallback for text mismatches). If no match — even
   with a high Zendesk-only count — **skip entirely**, don't fall back to Zendesk-only
   judgment. Currently this sheet only covers iPhone 18 Series, so SIREN is effectively
   iPhone-18-only in practice right now.
3. **Tag derivation**: defect tag matched by suffix-only regex (never guess the category
   prefix — e.g. `(SP)_레인보우현상` → tag prefix is `steinheil`, not `sp`). Device/product
   tags matched by exact-normalized-match first, then token-containment fuzzy fallback.
4. **Count**: live Zendesk search (`tags:a,b,c` comma-joined — repeating `tags:` fields is
   OR, not AND) + `1-3점` bad-review sheet rows joined by ASIN/SKU (not 모델명 text, which
   can differ between sheets). `total = zendesk_count + bad_review_count`.
5. **Trigger**: `total >= 3`. Alert posted as a **thread reply** on the original report's
   Chat thread (via `TicketQueue` lookup) — matching Zendesk ticket links + bad-review
   images, capped at 10 each.
6. **Dedup**: one alert per (defect_tag, product_tag) cluster, ever — later tickets in an
   already-triggered cluster log quietly without re-alerting.
7. Every check (triggered or not) is logged to both `state/siren_log.json` and the
   `SIREN_Log` sheet tab.

---

## Other things worth a look (lower confidence than the greeting bug — flagging, not fixing)

- **The `/revision` feedback check isn't actually being run every tick.** SKILL.md step 0
  says `check_feedback.py list` should run at the start of *every* tick. In practice today
  I only ran it when you asked directly, and it turned out two real corrections were
  sitting unapplied. Worth deciding: should every future tick's response explicitly show
  it ran this check, so a silent skip is visible?
- **SIREN's 신제품 라인업 scope means most defect terms currently never trigger anything**
  (Pixel 11, Galaxy S26, etc. tickets all skip regardless of cluster size) — this is
  working as designed per your explicit instruction, just flagging it stays a standing
  limitation until that sheet gets more product rows.
- **The `siren-report` Slides generator isn't wired into this alert flow yet** — a SIREN
  trigger still only posts a thread-reply with links/images, not an auto-built deck. Noted
  as future work in that skill's own SKILL.md.

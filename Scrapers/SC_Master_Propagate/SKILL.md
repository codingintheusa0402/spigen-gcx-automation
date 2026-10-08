---
name: sc-review-propagate
description: >-
  After a new `SC_yymmdd` (SC scraper) or `CaspiLM_yymmdd` (Caspi fetch) tab is
  added to the source review spreadsheet, funnel it into the master `SC` sheet
  (append minus header, dedupe by Review ID), propagate the new rows to the 8
  active product monitoring books via the `<Product> finalize` filter views
  (stamping Update 날짜, 키워드 =ai() and =dr() where each book expects them),
  restyle the Update 날짜 column (today = yellow+bold, older = plain), refresh
  the `tem` sheet, delete the older dated tabs (newest kept, only if all their
  rows are in `SC`), post the "Bad Review Monitoring Completed" cardsV2
  card (per-sheet Open buttons) to the GCX Chat webhook, and — as the LAST
  step, always, regardless of source — live-check today's rows flagged
  사진 유무=Y with an empty Image URL against the real Amazon review page and
  backfill the URL if a photo genuinely exists (Phase G). Trigger when the user says "propagate the SC sheet",
  "cascade SC_yymmdd / CaspiLM_yymmdd downstream", "run the SC review
  propagation", "append to SC sheet and distribute", "do the job" after a
  scrape, or any close paraphrase — and offer it right after any SC scraper run
  or CaspiLM fetch completes.
metadata:
  category: automation
  locale: ko-KR
  phase: v1.0.0-live
---

# sc-review-propagate

Turns a freshly-added `SC_yymmdd` / `CaspiLM_yymmdd` tab on the source
spreadsheet `1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw` into distributed rows
across the product monitoring books the Korean CX team works from. Every rule
below was given by the user 2026-09-10 → 2026-09-11 and verified live.

## How it runs (what kind of thing this is)

- **This is a Claude Code skill** (this file) that drives a **Python script**:
  `~/Desktop/GCX/Scrapers/SC_Master_Propagate/propagate.py` — plain Python 3 +
  `googleapiclient` (Google **Sheets API v4** and **Drive API v3** over REST).
- **Not MCP, not Apps Script.** No MCP connector is involved in propagation.
  The legacy Apps Script pipeline (`GAS_ReviewAutomation/MasterTrigger/Master.js`
  `dailyJob()`, fed by Apify) still runs on its own; it writes the same
  destination sheets but is a separate, untouched path — Review-ID dedup at each
  destination is what keeps the two from duplicating rows.
- Credentials: `~/.config/gws_shim/token.json` (OAuth user token,
  kjw@spigen.com, `drive` scope — see memory `gws_shim_sheets_token` /
  `gws_shim_gcp_project`). The script refreshes and rewrites the token itself.
- Upstream producers of the input tab are separate: the SC scraper
  (`Scrapers/SC_Review_Scraper/scrape_sc_reviews.py`, Python + Playwright) or
  the CaspiLM fetch (MCP `run_query`, see memory `caspilm_review_fetch_workflow`).
- Canonical copy of this file lives in the repo next to the script;
  `~/.claude/skills/sc-review-propagate/SKILL.md` is a symlink to it.

## Run order (the whole job)

```bash
cd ~/Desktop/GCX/Scrapers/SC_Master_Propagate
python3 propagate.py --new-sheet SC_260911              # Phase A dry-run → counts
python3 propagate.py --new-sheet SC_260911 --commit     # Phase A
python3 propagate.py --all-products                     # Phase B worklist (dry-run)
python3 propagate.py --product GlxZ8 --commit           # Phase B+D, one product at a time
python3 propagate.py --product Pixel11 --commit         #   append-at-bottom books first…
python3 propagate.py --product iPh18 --commit
python3 propagate.py --product SDA --commit
python3 propagate.py --product Auto_Acc --commit
python3 propagate.py --product Power_Acc --commit
python3 propagate.py --product 전략폰 --commit
python3 propagate.py --product 유지훈P --commit          #   …insert-at-row-2 book last
python3 propagate.py --refresh-tem --commit             # Phase C
python3 propagate.py --all-products                     # must end at 0 pending everywhere
python3 propagate.py --cleanup --notify --new-sheet SC_260911 --commit   # Phase F + E (see below)
# Phase G (agent-run, no CLI flag — see below): live-check + backfill Image URL
# on every sheet just touched, for rows dated today with 사진 유무=Y and empty
# Image URL. ALWAYS run this last, every time, whichever source fed the run.
```

Everything is dry-run without `--commit`. `--finish --product X --commit`
re-runs only the post-paste steps (`=dr()` on today's un-classified rows + the
date restyle) and is idempotent — use it if a run died half-way. After each
paste, read back `A{first-1}:K{first}` and confirm the row above the block is
unchanged (see Lessons). To force a test/backfill date on the scraper:
`SC_SCRAPER_RUN_DATE=260911 python3 scrape_sc_reviews.py` → writes `SC_260911`.

First full live run 2026-09-11 (`SC_260911`): 1,973 scraped → 319 new into `SC`
→ 68 rows pasted (GlxZ8 32 · Pixel11 16 · 유지훈P 16 · SDA 1 · Auto_Acc 2 ·
Power_Acc 1 · 전략폰 0). A quiet day can legitimately yield 0 at every stage.

## Source sheets

**`SC` master** — real name `SC`, **gid `444769313`** (NOT the dated
`SC_2609xx` tabs). Data cols `A:N` = `ASIN | Created 날짜 | 사진 유무 | Reviewer |
Review Ratings | Review Title | 본문 | 국가 | Review Link | Image URL | Review ID
| Order ID | Product Rating | Ratings Count`. Cols `O:X` (`Device`, `GlxZ8 Tab`,
`Pixel11 Tab`, `유지훈P Tab`, `AutoAcc Tab`, `PowerAcc Tab`, `SDA Tab`,
`전략폰 Tab`, `GlxS26 Tab`, `iPh18 Tab`) are `MAP`/`ARRAYFORMULA` helpers over
open-ended ranges — **never write to `O:X`**, they auto-spill onto appended rows.

**`tem`** (gid `902775794`), header row 1, Review IDs row 2+:
`A SDA | B iPh17 | C Auto Acc | D 전략폰 | E Power_Acc | F 유지훈P | G Pixel 10a |
H Glx26 | I iPh17e | J GlxZ8 | K Pixel11 | L iPh18`. Each `<Product> Tab` formula
on `SC` is `COUNTIF(tem!<col>2:<col>, K2:K) > 0 ? "Updated" : "Update Needed"`
(`GlxZ8→J`, `Pixel11→K`, `유지훈P→F`, `AutoAcc→C`, `PowerAcc→E`, `SDA→A`,
`전략폰→D`, `GlxS26→H`, `iPh18→L`).
- **Cols A-E** are `={"hdr"; IMPORTRANGE(<book>, "…J2:J")}` — self-updating,
  **never write to A-E** (IMPORTRANGE is only used on the small books; the big
  ones hit mass-fetch errors).
- **Cols F-L** are plain value lists that **Phase C rewrites every run** from
  each product's `1-5점` Review-ID col **K** (유지훈P has no `1-5점` → its `1-3점`
  col K). Source books: 유지훈P `1dlY6q8t…`, Pixel 10a `1BpeGq5g…`, Glx26
  `1fpv9TEDP…`, iPh17e `16xRJHH7…`, GlxZ8 `19Ohswgl…`, Pixel11 `12I6z_FFm…`,
  iPh18 `1aYxZRm7…` (added 2026-09-21).

**Filter views on `SC`** — the per-product selection logic. Read them live
(`spreadsheets.get(fields=sheets.filterViews)`), never hardcode. As of
2026-09-10 (`Device` values are comma-joined, e.g. `유지훈P, GlxZ8`):

| View | Criteria (0-based col → rule) |
|---|---|
| `GlxZ8 finalize` | col1 `Created 날짜` DATE_AFTER 2026-07-31 · col14 `Device` ∋ `GlxZ8` · col15 `GlxZ8 Tab` = `Update Needed` |
| `Pixel11 finalize` | col14 `Device` ∋ `Pixel11` · col16 `Pixel11 Tab` = `Update Needed` |
| `유지훈P_finalize` | col4 `Review Ratings` ∈ {1,2,3} · col14 `Device` ∋ `유지훈P` · col17 `유지훈P Tab` = `Update Needed` |
| `SDA finalize` | col4 ∈ {1,2,3} · col7 `국가` ∉ {US,IN} · col14 `Device` = `SDA` · col20 `SDA Tab` = `Update Needed` |
| `AutoAcc finalize` | col4 ∈ {1,2,3} · col7 ∉ {US} · col14 `Device` = `AutoAcc` · col18 `AutoAcc Tab` = `Update Needed` |
| `PowerAcc finalize` | col4 ∈ {1,2,3} · col7 ∉ {US} · col14 `Device` = `PowerAcc` · col19 `PowerAcc Tab` = `Update Needed` |
| `전략폰 finalize` | col4 ∈ {1,2,3} · col7 ∉ {US,JP,IT} · col14 `Device` ∋ `전략폰` · col21 `전략폰 Tab` = `Update Needed` *(fixed via API 2026-09-10; it used to hide only blank)* |
| `GlxS26 finalize` | inactive — Galaxy S26 is past its monitoring period |
| `iPh18 finalize` | col1 `Created 날짜` DATE_AFTER 2026-09-17 (user-narrowed 2026-09-21, was 2026-09-10) · col14 `Device` ∋ `iPh18` · col23 `iPh18 Tab` has no active condition (harmless — Phase B's own dest-dedup still prevents re-pasting) |

## Destination books (verified 2026-09-10/11)

| Book | Paste into | Paste cols | Review ID | Also stamp | `=dr()` |
|---|---|---|---|---|---|
| **GlxZ8** `19Ohswgl…`, **Pixel11** `12I6z_FFm…`, **iPh18** `1aYxZRm7…` | `1-5점` (append at bottom) | `A:K` | col K, **pasted raw** (the CX team relies on it) | `L` 키워드 `=ai(…,G{n})` · `P` Update 날짜 | on `1-3점` col **M** (see Phase D) |
| **유지훈P** `1dlY6q8t…` | `1-3점` only — **insert at row 2** | `A:J` | col K, **auto** (REGEXEXTRACT of Review Link) — not pasted | `M` Update 날짜 | col **L** on the inserted rows |
| **SDA** `1sxapIqJg…`, **Auto_Acc** `1mEYb1b9…`, **Power_Acc** `1QC8Is6U…`, **전략폰** `1yo8CbLh…` | `1-3점` only (append at bottom) | `A:I` | col J, **auto** from Review Url (I) — not pasted | `N`/`M`/`M`/`M` Update 날짜 (`Exported Date` on Auto_Acc) | **none** — CX agents type 인입사유 by hand |
| ~~Glx26 `1fpv9TEDP…`~~ | inactive (config kept, skipped by `--all-products`) | | | | |

Everything past the paste boundary other than the two stamps is a row-1 array
literal (SKU/기종명/모델명/대분류/…, `Customer Order ID` MAP, `국가`, 인입사유
mirrors) or agent-typed — never written. `Review Ratings` is written as a
number (the `1-3점` FILTER compares `E<=3` numerically).

## Procedure

### Phase A — funnel the new tab into `SC` (`--new-sheet <tab>`)
1. Read `'<tab>'!A:N`, drop the header. Read `'SC'!K2:K` (Review-ID column only
   — the full `A:N` read of ~14k rows timed out).
2. Append only rows whose Review ID is not already in `SC` (= "append then
   dedupe keeping the first occurrence").
3. Extend any `* finalize` filter-view `endRowIndex` the new row count exceeds
   (over-provisioned to 17,818 as of 2026-09-11, so rarely needed).

### Phase B — select + paste, per product (`--product X --commit`)
1. Apply the product's live filter-view criteria to `'SC'!A2:W` → candidates.
2. Drop candidates whose Review ID is already in the destination sheet.
3. Paste per the table above (RAW), then stamp `Update 날짜` (today, KST) and,
   on `1-5점` books, the per-row `키워드` formula
   `=ai("briefly summarize input which is customer's amazon product review of
   our(Spigen) product. max 10 words in english only",G{n})` (reads back as
   literal text via the API until the sheet is opened and Gemini evaluates it —
   normal).
4. **유지훈P insert-at-top**: snapshot row-1 formulas (FORMULA render) →
   `insertDimension` rows 2..N+1 with `inheritFromBefore: false` (format comes
   from the old row 2) → **rewrite row 1 from the snapshot** (inserting at row 2
   shifts `$I$2:$I` refs to `$I$3` — the breakage the user warned about) → write.
5. Run the date restyle (below) on the sheet — and on its `1-3점` for has15 books.

### Standing rule — date-cell styling (user, 2026-09-11)
In the `Update 날짜` / `Exported Date` column of every sheet touched: cells dated
**today** → yellow fill + bold; **every other date → no fill, not bold**.
`restyle_update_dates()` clears rows 2..last then re-applies on today's rows.

### Phase C — refresh `tem` (`--refresh-tem --commit`)
Rewrite cols **F-K only** from the source books above; clear residue below
(to row 6000); **never touch A-E**. Run it **every time**, even on a 0-new run —
it is the full `tem`↔reality sync that turns stale `Update Needed` rows back to
`Updated` so the filter views stop over-selecting.

### Phase D — `=dr()` on 인입사유(AI) — GlxZ8 / Pixel11 / iPh18 / 유지훈P only
Set `=dr(G{n}, S{n})` (본문, 대분류) on rows whose `Update 날짜` == today and
whose 인입사유(AI) is empty — the same rule `Master.js` uses, hence idempotent.
- **GlxZ8 / Pixel11 / iPh18**: `1-3점` A:L is `=FILTER('1-5점'!A2:L, '1-5점'!E2:E<=3)`
  **anchored at A2** (a row-1 formula scan misses it). **Never paste into
  `1-3점`** — the new ≤3★ rows appear by themselves, in `1-5점` order, and col O
  (`Update 날짜`) is a matching FILTER. Only col **M** gets the per-row `=dr()`.
  Cols M/N are static per-row cells: deleting or re-rating an existing `1-5점`
  row shifts everything below it in `1-3점` M/N by one — don't touch existing
  `1-5점` rows.
- **유지훈P**: col **L** on the freshly inserted rows 2..N+1.
- **SDA / Auto_Acc / Power_Acc / 전략폰**: skip (`dr_skip: True`).

### Phase F — delete older dated tabs (DEFAULT, user rule 2026-09-14)
`--cleanup [--commit]`. Once the newest `SC_yymmdd` / `CaspiLM_yymmdd` tab has
been distributed, the older tabs of each family are deleted (their rows already
live in master `SC`). Rules the script enforces:
- The **newest tab of each family is always kept** (families are matched by
  `^(SC|CaspiLM)_\d{6}`; `CaspiLM_BadReview_…`-style tabs are not touched).
- A tab is deleted **only if every Review ID in it is already in `SC`**. If any
  are missing it is kept and reported — run `--new-sheet <that tab> --commit`
  first (that is exactly what happened 2026-09-14: `SC_260907` and two
  CaspiLM tabs had 1/1/7 never-funneled rows; funnel → then cleanup).
- First live run 2026-09-14 removed 8 tabs (`SC_260831`, `SC_260907`,
  `SC_260908`–`11`, `CaspiLM_260827_asof0830`, `CaspiLM_260905`).

### Phase E — completion notice to Google Chat (user rule, 2026-09-14; cardsV2 since the same day)
`--notify --new-sheet <tab> [--removed a,b] [--commit]` posts a **cardsV2 app
card** to the GCX Chat incoming webhook (`NOTIFY_WEBHOOK` in `propagate.py`;
the URL lives only in the script and in memory `sc_master_sheet_propagation`).
Run it together with `--cleanup` so the removed tabs are listed automatically.
Layout (verified against the cardsV2 notes in memory `gchat_cardsv2_schema_reference`):
- Header: **Bad Review Monitoring Completed** · subtitle `yyyy-mm-dd · SC scraper → SC master → 7 monitoring books`, Sheets logo.
- **Source** section: `SC_yymmdd` (reviews scraped) and master `SC` (rows total), each a `decoratedText` with an **Open** button (`openLink` → `…/edit#gid=<tab gid>`).
- **Monitoring sheets · +N rows today** section: one `decoratedText` per touched tab — topLabel = product, text = `<b>live Drive spreadsheet name</b> · tab`, bottomLabel = `+N rows added today` (N = rows whose `Update 날짜`/`Exported Date` is today, by any pipeline), **Open** button to that tab. Has15 books list `1-5점` and its `1-3점` mirror; the section total counts only the paste tabs.
- **Housekeeping**: `tem` refreshed, `Removed older tabs: …`.
- Spreadsheet names come from `drive.files.get` at send time — never hardcoded. `--notify` without `--commit` prints the JSON and does not send. First card sent 2026-09-14.

### Phase G — Image URL backfill (user rule 2026-09-29, MANDATORY LAST STEP)
**Always runs after Phase E, every single propagate run, regardless of whether
the source was `SC_yymmdd` (scraper) or `CaspiLM_yymmdd` (Caspi).** This is
**agent-run** (Claude driving `mcp__claude-in-chrome__*` tools) — not a
`propagate.py` flag — because it needs a real rendered browser page, which the
Sheets/Drive-API-only script can't do.

**Why it exists**: CaspiLM's `AMAZON_REVIEW_PHOTOS` table (and, separately, the
SC Scraper's own image-fetch pass) both under-report — many rows land in a
destination sheet flagged `사진 유무=Y` with an **empty** Image URL even though
the review genuinely has a customer photo. Verified repeatedly this session at
38–100% hit rates (see memory `caspilm_review_fetch_workflow` and the
`SC_vs_CaspiLM_ImageURL_*` comparison sheet) — this is a real, recurring gap,
not a one-off.

**Scope**: for every destination sheet this run's Phase B touched (i.e., every
active product's `1-5점` if it has one, else `1-3점`), find rows where
**`Update 날짜` == today AND `사진 유무` == 'Y' AND `Image URL` is empty**.
Resolve the `사진 유무` / `Image URL` / `Update 날짜` / `Review Link` / `Review ID`
column letters **dynamically per sheet via header lookup** (`find_col`-style,
exact match on header text) — never hardcode a column letter, since it differs
per book (`1-5점` A:V vs `1-3점` A:V vs the has-Image-URL-at-different-position
`SDA`/`Auto_Acc`/`Power_Acc`/`전략폰` layout — check the live header first).
Only today's rows — **not** the whole historical sheet (that's a separate,
occasional user-requested backlog sweep, e.g. the 2026-09-29 유지훈P `1-3점`
K2:K168 and iPh18 `1-5점` J536:J604 one-off cleanups — do those only when asked).

**On the GCX server (Linux / WSL `gcx-server`) there is no claude-in-chrome** — run the script
version instead of the steps below (same rules: today's rows only, header lookup, amazon.es skipped,
re-check before write, `|`-joined):
```bash
cd ~/Desktop/GCX/Scrapers/SC_Master_Propagate
python3 phase_g.py            # dry run — shows what it would fill
python3 phase_g.py --commit   # write
```
It opens a headed Chrome with profile `~/.chrome-phaseg-profile`. **Amazon review pages need a
signed-in customer account** (anonymous visits redirect to /ap/signin). If the summary shows
"sign-in or captcha page", tell the user to open the laptop desktop shortcut
**"Amazon logins (photo check)"** and sign in on each Amazon tab once, then re-run `--commit`.
Books without a 사진 유무 column (Auto_Acc, Power_Acc, 전략폰) are reported as "not checked" — same as the Mac flow.

**Procedure per candidate row**:
1. Read the row's `Review Link` (already in the sheet — don't reconstruct it).
2. Group all candidates across all touched sheets by domain (`www.amazon.<tld>`)
   so same-domain rows can run in one `browser_batch` call — **`browser_batch`
   refuses cross-domain navigation mid-batch**, so batch per domain, not globally.
3. For each: `navigate` to the Review Link, then run this exact extraction JS
   (proven this session, handles multi-image reviews and dedupes CDN resize
   variants of the same photo):
   ```js
   const rid='<REVIEW_ID>';
   const el=document.getElementById(rid)||document.querySelector('[data-hook="review"]');
   if(!el){JSON.stringify({found:false});}else{
     const imgs=Array.from(el.querySelectorAll('img')).map(i=>i.src)
       .filter(s=>s.includes('/images/I/')&&!s.includes('avatars')&&!s.includes('/sash/'));
     const map={};
     imgs.forEach(u=>{const fn=u.split('/').pop();const base=fn.split('.')[0];
       if(!map[base]||!u.includes('_SY500_')){map[base]=u;}});
     JSON.stringify({found:true, imgs:Object.values(map)});
   }
   ```
4. **`amazon.es` needs a customer sign-in** in this Chrome profile — it wasn't
   logged in as of 2026-09-23/29 (unlike `.com`/`.de`/`.it`/`.in`/`.co.uk`/`.fr`
   which all work directly). **Never log in yourself.** Skip `.es` rows, report
   them as "needs amazon.es login" in the summary, and if the user says they've
   authed, retry just those rows.
5. **Safety check immediately before writing**: re-read the row's Image URL +
   Review ID cells and confirm Image URL is still empty and Review ID still
   matches, in case something else wrote to the sheet in between. Then join
   found URLs with `|` (matches the SC-scraper/CaspiLM multi-image convention)
   and write only that cell — never touch other columns.
6. Report a short summary: candidates found, filled, and any skipped (login-
   walled domain, or genuinely no image found on the live page — expected to
   be rare/zero based on every check run so far).

## Lessons

- **2026-09-11 — off-by-one overwrote a live row.** `last_data_row` returned the
  last *occupied* row instead of the next free one; the first GlxZ8 paste
  overwrote `1-5점` row 1358 (a Naver/KR review whose Review ID is a
  `phinf.pstatic.net` URL). Fixed (+1). The row was recovered **from Drive
  revision history via the API**: `drive.revisions.list(fileId)` → last revision
  before the bad write → `GET https://docs.google.com/spreadsheets/d/{id}/export
  ?format=csv&gid={gid}&revision={revId}` with the gws_shim token as a Bearer
  header → parse CSV → write back (RAW A:K; USER_ENTERED for the `=ai()` formula
  and the date; copy the neighbour's `numberFormat` onto the date cell) →
  `--finish` + `--refresh-tem`. Always read back the row above a pasted block.
- `googleapiclient`'s 60 s default timed out on wide reads of `SC`; the client
  is built with `httplib2.Http(timeout=300)`.
- Some sheets in these books (connected dashboards) have no `gridProperties`;
  tolerate it.
- **유지훈P `1-3점` row-1 invariant:** the 10 array-literal formulas in A1:V1 hash
  (md5 of the FORMULA-rendered row) to `bf423486dc…` — identical before/after the
  2026-09-11 and 2026-09-14 inserts. Recompute after every insert-at-top run; a
  different hash means the `$X$2` refs shifted and row 1 must be rewritten.
- `last_data_row` scans A + the Review-ID col + the paste-through col (not just
  A/K): the J-col books have agent-typed rows with no ASIN or Review ID.

## Decisions log (all from the user)

1. Glx26 (`1fpv9TEDP…`, has 1-5점 + 1-3점) is past its monitoring period →
   `inactive`. 8 active products (added iPh18 2026-09-21).
2. No `=dr()` on SDA / Auto_Acc / Power_Acc / 전략폰 — agents type 인입사유.
3. `전략폰 finalize` must show only `Update Needed` (view fixed).
4. J-col books paste `A:I` only; 유지훈P `A:J` only. Everything past is
   auto/agent-typed.
5. `tem` A-E are IMPORTRANGE — leave alone; F-K rewritten from `1-5점` (or
   `1-3점` when no 1-5점) Review-ID col.
6. Update 날짜 / Exported Date: today's cells yellow+bold, older cells plain.
7. **Still open:** should `scrape_sc_reviews.py` call Phase A itself at the end
   of its upload, or does it stay a separate step run by Claude? (Currently
   separate.)
8. **Phase G (Image URL backfill) is now a permanent, mandatory last step**
   (2026-09-29) — runs after every Phase E, on every run, no matter the source.
   Not a `propagate.py` flag; agent-run via `mcp__claude-in-chrome__*`. See its
   own section above for the full procedure, the proven extraction JS, and the
   `amazon.es`-needs-login caveat.

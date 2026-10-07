# SC_Master_Propagate

`propagate.py` funnels a newly added `SC_yymmdd` (SC scraper) or `CaspiLM_yymmdd` (Caspi fetch) tab on the source review spreadsheet (`1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw`) into the master `SC` sheet, then distributes the new rows to the 8 active downstream product monitoring books. It also refreshes the `tem` sheet, deletes older dated tabs, and posts a "Bad Review Monitoring Completed" Google Chat card.

It is plain Python 3 with `googleapiclient` (Sheets API v4 / Drive API v3). It is not MCP and not Apps Script; the legacy `MasterTrigger/Master.js` Apify pipeline is a separate path.

The full procedure, schemas, filter-view criteria, decisions and lessons are in the Claude Code skill **[`SKILL.md`](SKILL.md) in this folder** (canonical, git-tracked). `~/.claude/skills/sc-review-propagate/SKILL.md` is a symlink to it, so Claude Code loads it as the `sc-review-propagate` skill ("propagate the SC sheet", "do the job" after a scrape). `README_SC_Master_Propagate.md` is the older short version of this file.

## Screenshots

![Phase E "Bad Review Monitoring Completed" card (built with dry-run, posted only to the private test room)](docs/notify_card.jpg)
*Phase E "Bad Review Monitoring Completed" card (built with dry-run, posted only to the private test room)*

## Pipeline

| Phase | Flag | What it does |
|---|---|---|
| **A** | `--new-sheet <tab>` | Appends the tab's data rows (no header) to master `SC` (14 cols A..N), dedupes `SC` by Review ID (col K, first occurrence kept), and extends the 8 `<Product> finalize` filter views to the new rows. |
| **B** | `--product <P>` / `--all-products` | Applies each product's `<Product> finalize` filter-view criteria to `SC`, drops Review IDs already in the destination, and pastes the survivors into the product's `1-5점` sheet (or `1-3점` where the book has only that). Stamps `Update 날짜`, the `키워드` `=ai()` formula where the book expects it, and restyles `Update 날짜` (today = yellow + bold, older = plain). |
| **C** | `--refresh-tem` | Rewrites `tem` cols F–L (유지훈P, Pixel 10a, Glx26, iPh17e, GlxZ8, Pixel11, iPh18) from each book's Review-ID column. Cols A–E are `IMPORTRANGE` and are never touched. Run it every time. |
| **D** | (part of B, or `--finish`) | Sets 인입사유(AI) = `=dr(<본문>, <대분류>)` on today's new `1-3점` rows. GlxZ8 / Pixel11 / iPh18 / 유지훈P only; SDA / Auto_Acc / Power_Acc / 전략폰 are `dr_skip` (typed by hand). |
| **E** | `--notify` | Builds the cardsV2 card "Bad Review Monitoring Completed": Source (today's scrape tab + master `SC` total), one row per monitoring tab with "+N rows added today" and an **Open** button, and a Housekeeping footer (tem refresh, removed tabs). Posts it to the GCX room webhook (`NOTIFY_WEBHOOK` in `propagate.py`). |
| **F** | `--cleanup` | Deletes older `SC_yymmdd` / `CaspiLM_yymmdd` tabs. The newest of each family is always kept, and a tab is deleted only if all its Review IDs are already in `SC`. |
| **G** | — (agent-run) | Last step, run by the skill rather than the script: for today's rows with `사진 유무 = Y` and an empty `Image URL`, check the live Amazon review page and backfill the URL if a photo really exists. See `SKILL.md`. |

**Safety: everything is a dry run unless you pass `--commit`.** A dry run reads the sheets and prints the worklist (or, for `--notify`, the card JSON) and writes nothing.

Rows are only funnelled or pasted if their Review ID matches `^R[A-Z0-9]{6,20}$`. This guard was added 2026-09-15, after a stray header row in a source tab was pasted into every book.

## Usage

```bash
# Phase A: append the new tab into the master SC sheet (dry run, then live)
python3 propagate.py --new-sheet SC_260910
python3 propagate.py --new-sheet SC_260910 --commit

# Phase B/C/D for one product, dry run (prints the worklist)
python3 propagate.py --product GlxZ8

# all active products, dry run only
python3 propagate.py --all-products

# Phase B/D live for one product (A:<boundary> + Update 날짜 + 키워드 =ai()
# + =dr() on today's 1-3점 rows + date restyle)
python3 propagate.py --product GlxZ8 --commit

# re-run only the post-paste steps (idempotent): =dr() on today's
# un-classified 1-3점 rows + Update 날짜 restyle
python3 propagate.py --finish --product GlxZ8 --commit

# Phase C: rewrite tem cols F-L
python3 propagate.py --refresh-tem --commit

# Phase F + E: delete older dated tabs, then post the completion card
python3 propagate.py --cleanup --notify --new-sheet SC_260914 --commit

# Preview the completion card without sending it (dry run prints the JSON)
python3 propagate.py --notify --new-sheet SC_260914
```

Run products one at a time. After each paste, read back the row just above the pasted block (see the skill's 2026-09-11 lesson). When `--notify` is set, `--new-sheet` only labels the card; Phase A is skipped.

## Products

All product-specific settings are in the `PRODUCTS` dict at the top of `propagate.py`: destination book id, paste sheet, `1-3점` mirror, Review-ID column, whether the raw Review ID is pasted, insert-at-top vs append, `tem` column and `=dr()` target.

| Product | Paste sheet | 1-3점 mirror | `=dr()` |
|---|---|---|---|
| GlxZ8 | 1-5점 | 1-3점 | yes |
| Pixel11 | 1-5점 | 1-3점 | yes |
| iPh18 | 1-5점 | 1-3점 | yes |
| 유지훈P | 1-3점 (insert at top) | — | yes |
| SDA | 1-3점 | — | skip |
| Auto_Acc | 1-3점 | — | skip |
| Power_Acc | 1-3점 | — | skip |
| 전략폰 | 1-3점 | — | skip |
| Glx26 | — | — | inactive |

## Recent changes

- 2026-10-01: Phase C now refreshes `tem` col L (iPh18), which it had never updated.
- 2026-09-15: Phase A and B reject header-like or malformed rows (Review-ID regex).
- 2026-09-14: Phase F cleanup (on by default in the skill) and the Phase E cardsV2 completion card were added. The `last_data_row` scan was widened.
- 2026-09-11: Live `--commit` for Phase B/D, `--finish`, and the `Update 날짜` restyle were added. The token path became portable (`$GWS_SHIM_TOKEN`) so the script also runs on the 24/7 server.

## Credentials

The token is `$GWS_SHIM_TOKEN` if set, else `~/.config/gws_shim/token.json` (kjw@spigen.com, `drive` scope). The script refreshes the token and rewrites the file on every run, including dry runs. The Chat webhook URL is a constant in `propagate.py`; don't copy it into docs.

## Dependencies

```bash
pip install google-api-python-client google-auth google-auth-httplib2 httplib2 requests
```

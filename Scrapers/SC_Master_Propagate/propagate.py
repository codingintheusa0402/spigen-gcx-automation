#!/usr/bin/env python3
"""
SC master-sheet propagation pipeline.

WHAT THIS DOES (established 2026-09-10, see skill `sc-review-propagate`):

  Phase A  After a new `SC_yymmdd` (SC scraper) or `CaspiLM_yymmdd` (Caspi fetch)
           tab is added to the SOURCE spreadsheet, append its data rows
           (WITHOUT header) to the master `SC` sheet, then dedupe the whole
           `SC` sheet by `Review ID` (col K) keeping the FIRST occurrence.
           Also extend the 8 `* finalize` filter views' row ranges to cover
           the new rows.

  Phase B  For each of the 8 downstream products, read its `<Product> finalize`
           filter view on the `SC` sheet, apply that view's criteria to the
           deduped `SC` rows, drop rows whose Review ID is already in the
           destination sheet, and paste the survivors into the product's
           monitoring sheet (1-5점 where the book has both 1-5/1-3, else 1-3점).

  Phase C  Rewrite `tem` sheet cols F-K (유지훈P, Pixel 10a, Glx26, iPh17e,
           GlxZ8, Pixel11) from each product's `1-5점` Review-ID col K (유지훈P
           uses its `1-3점` col K — no 1-5점). This makes the `<Product> Tab`
           helper formula on `SC` read true. Cols A-E (SDA, iPh17, Auto Acc,
           전략폰, Power_Acc) are `IMPORTRANGE` formulas and are NEVER touched.
           Run it every time (`--refresh-tem`), not only when rows were added.

  Phase D  On the 1-3점 sheet, set the 인입사유(AI) column for the newly added
           rows to `=dr(<본문col><n>, <대분류col><n>)`.
           GlxZ8 / Pixel11 / 유지훈P ONLY. SDA / Auto_Acc / Power_Acc / 전략폰
           skip Phase D (`dr_skip`) — their agents type 인입사유 by hand.

SAFETY: dry-run is the DEFAULT. Nothing is written without `--commit`.
Run ONE product at a time on the first live run and eyeball the sheet after
each: `--product GlxZ8 --commit`.

Credentials: ~/.config/gws_shim/token.json  (kjw@spigen.com, drive scope).
See memory gws_shim_sheets_token / gws_shim_gcp_project.
"""

import argparse
import datetime
import json
import os
import sys

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Resolved rather than hardcoded so the same file runs on the Mac and on the
# 24/7 Windows server. GWS_SHIM_TOKEN wins if set; otherwise the usual
# ~/.config/gws_shim/token.json, which expanduser() resolves correctly on both.
TOKEN = os.environ.get(
    "GWS_SHIM_TOKEN",
    os.path.expanduser("~/.config/gws_shim/token.json"),
)
SRC = "1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw"
SC_SHEET = "SC"
SC_SHEET_GID = 444769313
TEM_SHEET = "tem"

# SC / SC_yymmdd / CaspiLM_yymmdd canonical 14-col data layout (A..N):
#  A ASIN | B Created 날짜 | C 사진 유무 | D Reviewer | E Review Ratings |
#  F Review Title | G 본문 | H 국가 | I Review Link | J Image URL |
#  K Review ID | L Order ID | M Product Rating | N Ratings Count
# O..W on the SC sheet are ARRAYFORMULA/MAP helper columns (Device, <Product> Tab)
# that auto-spill down open-ended ranges — never write to O..W.
SC_DATA_COLS = 14
SC_REVIEW_ID_IDX = 10  # 0-based col K within the A..N block

# tem sheet columns (1-based), header row 1, Review IDs from row 2:
#  A SDA | B iPh17 | C Auto Acc | D 전략폰 | E Power_Acc | F 유지훈P |
#  G Pixel 10a | H Glx26 | I iPh17e | J GlxZ8 | K Pixel11
#
# Cols A-E are `={"hdr"; IMPORTRANGE(<destbook>, "…J2:J")}` formulas that
# self-update — NEVER write to A-E. (IMPORTRANGE is only used for the small
# books; the others have too many rows and would hit mass-fetch errors.)
#
# Cols F-K are plain value lists that Phase C must rewrite every run from the
# product's `1-5점` Review-ID column K (유지훈P has no 1-5점 → its `1-3점` col K).
# Verified 2026-09-10: all six sources use Review-ID col K.
TEM_REFRESH = {  # tem col letter -> (source book id, source sheet, review-id col)
    "F": ("1dlY6q8trbVMVJAjw_OUoxp1cguA2oTB8WlPhHR01xIw", "1-3점", "K"),  # 유지훈P
    "G": ("1BpeGq5gIr4tNsPZmnHr19NNY6pQ6sb2_H-v3V9-It4E", "1-5점", "K"),  # Pixel 10a
    "H": ("1fpv9TEDPGR8D6QRRc0ll-WzF7sOkfxe9UNBCmdBSE9g", "1-5점", "K"),  # Glx26 (inactive product, tem still kept fresh)
    "I": ("16xRJHH7Ynii4erNOn_905ST4CZs6OLpOYTof4uqsGsQ", "1-5점", "K"),  # iPh17e
    "J": ("19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4", "1-5점", "K"),  # GlxZ8
    "K": ("12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI", "1-5점", "K"),  # Pixel11
}
TEM_CLEAR_TO_ROW = 6000  # nuke any residue below the refreshed data

# ─────────────────────────────────────────────────────────────────────────────
# Per-product config. Verified 2026-09-10 against live sheet headers/filter views
# unless marked CONFIRM.
# ─────────────────────────────────────────────────────────────────────────────
PRODUCTS = {
    "GlxZ8": {
        "filter_view": "GlxZ8 finalize",          # FV 721716446
        "dest_id": "19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4",
        "dest_sheet": "1-5점",                     # book has 1-5점 AND 1-3점 → paste into 1-5점
        "dest_review_id_col": "K",
        "paste_review_id": True,                   # raw Review ID value IS pasted (CX team relies on it)
        "paste_through_col": "K",                  # paste A..K
        "insert_at_top": False,
        "one_three_sheet": "1-3점",                # mirror/filter of 1-5점 — do NOT paste into it directly
        "dr_sheet": "1-3점",
        "dr_col_header_contains": "인입사유(AI)",   # → col M ("인입사유(AI)  Acc. 74.5%")
        "dr_body_header": "본문",                  # → col G
        "dr_category_header": "대분류",            # → col S
    },
    "Pixel11": {
        "filter_view": "Pixel11 finalize",        # FV 529212369
        "dest_id": "12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI",
        "dest_sheet": "1-5점",
        "dest_review_id_col": "K",
        "paste_review_id": True,
        "paste_through_col": "K",
        "insert_at_top": False,
        "one_three_sheet": "1-3점",
        "dr_sheet": "1-3점",
        "dr_col_header_contains": "인입사유(AI)",   # col M ("인입사유(AI)  Acc. 83.9%")
        "dr_body_header": "본문",
        "dr_category_header": "대분류",
    },
    "Glx26": {
        # INACTIVE 2026-09-10: Galaxy S26 is past its monitoring period. Config
        # kept for reference only — skipped by --all-products; the `GlxS26
        # finalize` filter view is no longer maintained. Run explicitly with
        # --product Glx26 only if the user asks.
        "inactive": True,
        "filter_view": "GlxS26 finalize",         # FV 302123587
        "dest_id": "1fpv9TEDPGR8D6QRRc0ll-WzF7sOkfxe9UNBCmdBSE9g",  # confirmed 2026-09-10 (1-5점 + 1-3점)
        "dest_sheet": "1-5점",
        "dest_review_id_col": "K",
        "paste_review_id": True,
        "paste_through_col": "K",
        "insert_at_top": False,
        "one_three_sheet": "1-3점",
        "dr_sheet": "1-3점",
        "dr_col_header_contains": "인입사유(AI)",   # col M ("인입사유(AI)  Acc. 80.9%")
        "dr_body_header": "본문",
        "dr_category_header": "대분류",
    },
    "유지훈P": {
        "filter_view": "유지훈P_finalize",         # FV 216742354  (note underscore, no space)
        "dest_id": "1dlY6q8trbVMVJAjw_OUoxp1cguA2oTB8WlPhHR01xIw",
        "dest_sheet": "1-3점",                     # book has ONLY 1-3점 (and an unrelated "1~5점")
        "dest_review_id_col": "K",
        "paste_review_id": False,                  # K is auto-filled from Review Link (I) by a formula
        "paste_through_col": "J",                  # paste A..J only (through Image URL)
        "insert_at_top": True,                     # insert blank rows at row 2 each run, then fill; keep row-1 arrayformulas intact
        "one_three_sheet": None,
        "dr_sheet": "1-3점",
        "dr_col_header_contains": "인입사유(AI)",   # → col L ("인입사유(AI)")
        "dr_body_header": "본문",                  # → col G
        "dr_category_header": "대분류",            # → col S
    },
    # ── SDA / Auto_Acc / Power_Acc / 전략폰 ──────────────────────────────────
    # NO =dr() (confirmed 2026-09-10): CX agents type the 인입사유 column by hand
    # on these four. Phase D is skipped — `dr_skip: True`.
    "SDA": {
        "filter_view": "SDA finalize",            # FV 1125062509
        "dest_id": "1sxapIqJgXcJdeqyCf9bAxCNXrVMsVjsZE9QWPwEm0R4",
        "dest_sheet": "1-3점",                     # ONLY 1-3점
        "dest_review_id_col": "J",
        "paste_review_id": False,                  # J is derived from Review Url (I) by a formula
        "paste_through_col": "I",                  # paste A..I ONLY. J = Review ID (auto from I).
        # Everything from J onward is agent-typed or arrayformula — never pasted.
        "insert_at_top": False,
        "one_three_sheet": None,
        "dr_skip": True,
    },
    "Auto_Acc": {
        "filter_view": "AutoAcc finalize",        # FV 1131952515
        "dest_id": "1mEYb1b92D6BIOaSYkAnMit6THuw5ewtymhA-mSIVDfs",
        "dest_sheet": "1-3점",
        "dest_review_id_col": "J",
        "paste_review_id": False,
        "paste_through_col": "I",                  # paste A..I ONLY (J onward = auto / agent-typed)
        "insert_at_top": False,
        "one_three_sheet": None,
        "dr_skip": True,
    },
    "Power_Acc": {
        "filter_view": "PowerAcc finalize",       # FV 1414549791
        "dest_id": "1QC8Is6UvTnFXaOeXviKM_331i3Fo_CBIYx80VS696LI",
        "dest_sheet": "1-3점",
        "dest_review_id_col": "J",
        "paste_review_id": False,
        "paste_through_col": "I",                  # paste A..I ONLY (J onward = auto / agent-typed)
        "insert_at_top": False,
        "one_three_sheet": None,
        "dr_skip": True,
    },
    "전략폰": {
        "filter_view": "전략폰 finalize",          # FV 1853891342
        "dest_id": "1yo8CbLhJkuxrf3eXbAqZCb6qBejZhSR3YOt7nFv97fw",
        "dest_sheet": "1-3점",
        "dest_review_id_col": "J",
        "paste_review_id": False,
        "paste_through_col": "I",                  # paste A..I ONLY (J onward = auto / agent-typed)
        "insert_at_top": False,
        "one_three_sheet": None,
        "dr_skip": True,
    },
}

# `SC` sheet helper columns O..W (0-based within A..W) → the tem column each
# `<Product> Tab` formula compares against, for reference:
#   O(14) Device | P(15) GlxZ8 Tab→tem J | Q(16) Pixel11 Tab→tem K |
#   R(17) 유지훈P Tab→tem F | S(18) AutoAcc Tab→tem C | T(19) PowerAcc Tab→tem E |
#   U(20) SDA Tab→tem A | V(21) 전략폰 Tab→tem D | W(22) GlxS26 Tab→tem H


def a1_col_to_idx(letter):
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx  # 1-based


def idx_to_a1_col(idx):
    s = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


def get_service():
    with open(TOKEN) as f:
        info = json.load(f)
    creds = Credentials.from_authorized_user_info(info)
    creds.refresh(Request())
    with open(TOKEN, "w") as f:
        f.write(creds.to_json())
    # googleapiclient's default per-request timeout is 60s; the master SC sheet
    # is ~14k rows and full-width reads have hit "read operation timed out".
    import httplib2
    from google_auth_httplib2 import AuthorizedHttp
    http = AuthorizedHttp(creds, http=httplib2.Http(timeout=300))
    return build("sheets", "v4", http=http)


def load_filter_views(svc):
    meta = svc.spreadsheets().get(
        spreadsheetId=SRC,
        fields="sheets(properties(sheetId,title),filterViews(filterViewId,title,range,criteria))",
    ).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == SC_SHEET:
            return {fv["title"]: fv for fv in s.get("filterViews", [])}
    raise SystemExit("SC sheet not found on source spreadsheet")


def row_passes(row, criteria):
    """row: list of cell strings (A.. order). criteria: filterView['criteria']
    dict keyed by 0-based column index string."""
    for col_s, crit in criteria.items():
        col = int(col_s)
        val = row[col] if col < len(row) else ""
        hidden = crit.get("hiddenValues")
        if hidden is not None and val in hidden:
            return False
        cond = crit.get("condition")
        if cond:
            t = cond["type"]
            cvals = [v.get("userEnteredValue", "") for v in cond.get("values", [])]
            if t == "TEXT_CONTAINS":
                if cvals and cvals[0] not in val:
                    return False
            elif t == "TEXT_NOT_CONTAINS":
                if cvals and cvals[0] in val:
                    return False
            elif t in ("DATE_AFTER", "DATE_ON_OR_AFTER", "DATE_BEFORE",
                       "DATE_ON_OR_BEFORE", "DATE_EQ"):
                try:
                    d = datetime.date.fromisoformat(val[:10])
                    ref = datetime.date.fromisoformat(cvals[0][:10])
                except ValueError:
                    return False
                if t == "DATE_AFTER" and not d > ref:
                    return False
                if t == "DATE_ON_OR_AFTER" and not d >= ref:
                    return False
                if t == "DATE_BEFORE" and not d < ref:
                    return False
                if t == "DATE_ON_OR_BEFORE" and not d <= ref:
                    return False
                if t == "DATE_EQ" and not d == ref:
                    return False
            else:
                raise SystemExit(f"Unhandled filter condition type {t!r} — extend row_passes()")
    return True


def col_values(svc, sid, sheet, col_letter, start_row=2):
    rng = f"'{sheet}'!{col_letter}{start_row}:{col_letter}"
    r = svc.spreadsheets().values().get(spreadsheetId=sid, range=rng).execute()
    return [x[0] if x else "" for x in r.get("values", [])]


def header_row(svc, sid, sheet):
    r = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{sheet}'!1:1").execute()
    return r.get("values", [[]])[0] if r.get("values") else []


def find_col(headers, want_exact=None, contains=None):
    for i, h in enumerate(headers):
        if want_exact is not None and h.strip() == want_exact:
            return i + 1
        if contains is not None and contains in h:
            return i + 1
    return None


# ─────────────────────────────────────────────────────────────────────────────
def phase_a_append_and_dedupe(svc, new_sheet, dry_run=True):
    """Append `new_sheet` data rows (no header) to `SC`, dedupe SC by Review ID."""
    src_vals = svc.spreadsheets().values().get(
        spreadsheetId=SRC, range=f"'{new_sheet}'!A:N"
    ).execute().get("values", [])
    if not src_vals:
        raise SystemExit(f"{new_sheet} is empty")
    body = src_vals[1:]  # drop header
    body = [r + [""] * (SC_DATA_COLS - len(r)) for r in body if any(c.strip() for c in r)]

    # Only the Review-ID column is needed for dedup — reading the full A:N
    # block (14 cols incl. review text) was ~14x heavier and timed out.
    sc_body = svc.spreadsheets().values().get(
        spreadsheetId=SRC, range=f"'{SC_SHEET}'!K2:K"
    ).execute().get("values", [])
    seen = set(r[0] for r in sc_body if r and r[0])

    to_add = []
    dup_in_batch = 0
    for r in body:
        rid = r[SC_REVIEW_ID_IDX]
        if rid in seen:
            dup_in_batch += 1
            continue
        seen.add(rid)
        to_add.append(r)

    print(f"  {new_sheet}: {len(body)} rows in tab, "
          f"{dup_in_batch} already in SC (skipped), {len(to_add)} new to append")
    new_total = 1 + len(sc_body) + len(to_add)  # header + existing + new
    print(f"  SC row count: {1 + len(sc_body)} → {new_total}")

    if dry_run:
        print("  [dry-run] would append and extend 8 filter-view ranges to", new_total)
        return to_add

    if to_add:
        svc.spreadsheets().values().append(
            spreadsheetId=SRC, range=f"'{SC_SHEET}'!A1",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": to_add},
        ).execute()
    # extend filter view ranges
    fvs = load_filter_views(svc)
    reqs = []
    for fv in fvs.values():
        rng = dict(fv["range"])
        if rng.get("endRowIndex", 0) < new_total:
            rng["endRowIndex"] = new_total
            reqs.append({"updateFilterView": {
                "filter": {"filterViewId": fv["filterViewId"], "range": rng},
                "fields": "range",
            }})
    if reqs:
        svc.spreadsheets().batchUpdate(spreadsheetId=SRC, body={"requests": reqs}).execute()
        print(f"  extended {len(reqs)} filter-view ranges to row {new_total}")
    return to_add


# ─────────────────────────────────────────────────────────────────────────────
def phase_b_c_d_product(svc, product, dry_run=True):
    cfg = PRODUCTS[product]
    fvs = load_filter_views(svc)
    if cfg["filter_view"] not in fvs:
        raise SystemExit(f"filter view {cfg['filter_view']!r} not found on SC")
    crit = fvs[cfg["filter_view"]].get("criteria", {})

    sc_vals = svc.spreadsheets().values().get(
        spreadsheetId=SRC, range=f"'{SC_SHEET}'!A2:W"
    ).execute().get("values", [])
    cand = []
    for row in sc_vals:
        row = row + [""] * (23 - len(row))
        if row_passes(row, crit):
            cand.append(row)
    print(f"[{product}] filter view {cfg['filter_view']!r}: {len(cand)} rows pass criteria")

    dest_id, dest_sheet = cfg["dest_id"], cfg["dest_sheet"]
    existing_ids = set(col_values(svc, dest_id, dest_sheet, cfg["dest_review_id_col"]))
    new_rows = [r for r in cand if r[SC_REVIEW_ID_IDX] not in existing_ids]
    print(f"[{product}] {len(existing_ids)} already in {dest_sheet}; "
          f"{len(new_rows)} genuinely new")
    if not new_rows:
        return

    through = a1_col_to_idx(cfg["paste_through_col"])
    payload = []
    for r in new_rows:
        block = list(r[:through])
        if not cfg["paste_review_id"]:
            rid_i = a1_col_to_idx(cfg["dest_review_id_col"]) - 1
            if rid_i < len(block):
                block[rid_i] = ""
        payload.append(block)

    print(f"[{product}] will write {len(payload)} rows × {through} cols (A:{cfg['paste_through_col']}) into "
          f"'{dest_sheet}' ({'insert@row2' if cfg['insert_at_top'] else 'append@bottom'}); "
          f"paste_review_id={cfg['paste_review_id']}  "
          f"(everything past col {cfg['paste_through_col']} is auto-formula / agent-typed — never written)")

    dr_note = ("Phase D skipped (agents type 인입사유 by hand)"
               if cfg.get("dr_skip")
               else f"set 인입사유(AI) =dr() on '{cfg['dr_sheet']}'")
    if dry_run:
        print(f"[{product}] [dry-run] sample row:", payload[0][:12], "...")
        print(f"[{product}] [dry-run] would then {dr_note}. "
              "(Phase C tem refresh is a separate global step: --refresh-tem)")
        return

    phase_b_commit_product(svc, product, cfg, new_rows, payload)


KEYWORD_FORMULA = ('=ai("briefly summarize input which is customer\'s amazon product '
                   'review of our(Spigen) product. max 10 words in english only",G{n})')


def today_kst_iso():
    kst = datetime.timezone(datetime.timedelta(hours=9))
    return datetime.datetime.now(kst).strftime("%Y-%m-%d")


def sheet_props(svc, sid, title):
    meta = svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(sheetId,title,gridProperties))"
    ).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == title:
            s["properties"].setdefault("gridProperties", {"rowCount": 0})
            return s["properties"]
    raise SystemExit(f"sheet {title!r} not found in {sid}")


def ensure_rows(svc, sid, props, needed_last_row):
    """Grow the grid if the target rows don't exist yet (Master.js does +100)."""
    have = props["gridProperties"]["rowCount"]
    if needed_last_row <= have:
        return
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [{
        "appendDimension": {"sheetId": props["sheetId"], "dimension": "ROWS",
                            "length": needed_last_row - have + 100}}]}).execute()


def last_data_row(svc, sid, sheet, cols=("A", "K")):
    """Last row with content in any of `cols` (array-literal spills past the data
    render as "" and would fool values.append's table detection)."""
    best = 2  # first free row when the sheet has only a header
    for c in cols:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{sheet}'!{c}2:{c}").execute().get("values", [])
        last = 1
        for i, v in enumerate(vals):
            if v and str(v[0]).strip():
                last = i + 2          # sheet row number of that cell (range starts at row 2)
        best = max(best, last + 1)    # next FREE row (2026-09-11: an off-by-one here
    return best                       # overwrote a live row — keep the +1)


def numeric_rating(block, idx=4):
    """SC col E arrives as text ("5"); the 1-3점 FILTER does E<=3 numerically."""
    v = str(block[idx]).strip()
    if v.isdigit():
        block[idx] = int(v)
    return block


def phase_b_commit_product(svc, product, cfg, new_rows, payload):
    dest_id, dest_sheet = cfg["dest_id"], cfg["dest_sheet"]
    today = today_kst_iso()
    hdr = header_row(svc, dest_id, dest_sheet)
    upd_col = find_col(hdr, want_exact="Update 날짜") or find_col(hdr, want_exact="Exported Date")
    kw_col = find_col(hdr, want_exact="키워드 (AI 요약)")
    payload = [numeric_rating(list(b)) for b in payload]
    n = len(payload)
    width = a1_col_to_idx(cfg["paste_through_col"])
    props = sheet_props(svc, dest_id, dest_sheet)

    if cfg["insert_at_top"]:
        # Snapshot row-1 formulas: inserting at row 2 shifts their $X$2 refs to
        # $X$3 (the breakage the user warned about). Master.js rewrites them too.
        row1 = svc.spreadsheets().values().get(
            spreadsheetId=dest_id, range=f"'{dest_sheet}'!1:1",
            valueRenderOption="FORMULA").execute().get("values", [[]])[0]
        svc.spreadsheets().batchUpdate(spreadsheetId=dest_id, body={"requests": [{
            "insertDimension": {"range": {"sheetId": props["sheetId"], "dimension": "ROWS",
                                          "startIndex": 1, "endIndex": 1 + n},
                                "inheritFromBefore": False}}]}).execute()
        first = 2
        svc.spreadsheets().values().update(
            spreadsheetId=dest_id, range=f"'{dest_sheet}'!A1",
            valueInputOption="USER_ENTERED", body={"values": [row1]}).execute()
        print(f"[{product}] inserted {n} rows at row 2; row-1 formulas rewritten "
              f"({sum(1 for x in row1 if str(x).startswith('='))} formula cells)")
    else:
        first = last_data_row(svc, dest_id, dest_sheet)
        ensure_rows(svc, dest_id, props, first + n - 1)
    last = first + n - 1

    end_col = idx_to_a1_col(width)
    svc.spreadsheets().values().update(
        spreadsheetId=dest_id, range=f"'{dest_sheet}'!A{first}:{end_col}{last}",
        valueInputOption="RAW", body={"values": payload}).execute()
    print(f"[{product}] wrote {n} rows A:{end_col} → '{dest_sheet}' rows {first}-{last}")

    extra = []
    if upd_col:
        c = idx_to_a1_col(upd_col)
        extra.append({"range": f"'{dest_sheet}'!{c}{first}:{c}{last}",
                      "values": [[today]] * n})
    if kw_col:
        c = idx_to_a1_col(kw_col)
        extra.append({"range": f"'{dest_sheet}'!{c}{first}:{c}{last}",
                      "values": [[KEYWORD_FORMULA.format(n=r)] for r in range(first, last + 1)]})
    # 1-3점-only books with dr: stamp on the same rows
    if not cfg.get("dr_skip") and cfg.get("dr_sheet") == dest_sheet:
        ai_col = find_col(hdr, contains=cfg["dr_col_header_contains"])
        g = idx_to_a1_col(find_col(hdr, want_exact=cfg["dr_body_header"]))
        s = idx_to_a1_col(find_col(hdr, want_exact=cfg["dr_category_header"]))
        c = idx_to_a1_col(ai_col)
        extra.append({"range": f"'{dest_sheet}'!{c}{first}:{c}{last}",
                      "values": [[f"=dr({g}{r}, {s}{r})"] for r in range(first, last + 1)]})
    if extra:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=dest_id,
            body={"valueInputOption": "USER_ENTERED", "data": extra}).execute()
        print(f"[{product}] stamped: " + ", ".join(e["range"].split("!")[1] for e in extra)
              + f"  (Update 날짜={today})")

    # has15 books: 1-3점 A:L is `=FILTER('1-5점'!A2:L, E2:E<=3)` anchored at A2
    # (NOT row 1 — a row-1 formula scan misses it). Never paste into it; the
    # new <=3-star rows appear there by themselves. Just stamp =dr() on them.
    if cfg.get("one_three_sheet") and not cfg.get("dr_skip"):
        stamp_dr_on_filtered_13(svc, product, cfg, payload)
    restyle_update_dates(svc, dest_id, dest_sheet, today)
    if cfg.get("one_three_sheet"):
        restyle_update_dates(svc, dest_id, cfg["one_three_sheet"], today)


def stamp_dr_on_filtered_13(svc, product, cfg, payload=None, dry_run=False):
    """Same rule as Master.js: rows whose Update 날짜 == today and whose
    인입사유(AI) is empty get =dr(). Idempotent — safe to rerun (--finish)."""
    dest_id, s13 = cfg["dest_id"], cfg["one_three_sheet"] or cfg["dest_sheet"]
    hdr13 = header_row(svc, dest_id, s13)
    ai_i = find_col(hdr13, contains=cfg["dr_col_header_contains"])
    upd_i = find_col(hdr13, want_exact="Update 날짜") or find_col(hdr13, want_exact="Exported Date")
    ai13 = idx_to_a1_col(ai_i)
    g = idx_to_a1_col(find_col(hdr13, want_exact=cfg["dr_body_header"]))
    s = idx_to_a1_col(find_col(hdr13, want_exact=cfg["dr_category_header"]))
    today = today_kst_iso()
    y, m, d = (int(x) for x in today.split("-"))
    serial = (datetime.date(y, m, d) - datetime.date(1899, 12, 30)).days
    kor = f"{y}. {m}. {d}"
    grid = svc.spreadsheets().values().get(
        spreadsheetId=dest_id, range=f"'{s13}'!A2:{idx_to_a1_col(max(ai_i, upd_i))}",
        valueRenderOption="UNFORMATTED_VALUE").execute().get("values", [])
    rows = []
    for i, r in enumerate(grid):
        r = r + [""] * (max(ai_i, upd_i) - len(r))
        u = r[upd_i - 1]
        if (u == serial or str(u).strip() in (today, kor)) and str(r[ai_i - 1]).strip() == "":
            rows.append(i + 2)
    if not rows:
        print(f"[{product}] {s13}: no rows dated {today} with empty {hdr13[ai_i-1]!r}")
        return
    if dry_run:
        print(f"[{product}] {s13}: [dry-run] would set =dr({g},{s}) on {len(rows)} rows "
              f"({rows[0]}..{rows[-1]})")
        return
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=dest_id,
        body={"valueInputOption": "USER_ENTERED", "data": [
            {"range": f"'{s13}'!{ai13}{r}", "values": [[f"=dr({g}{r}, {s}{r})"]]}
            for r in rows]}).execute()
    print(f"[{product}] {s13}: =dr({g},{s}) set in col {ai13} on rows "
          f"{rows[0]}..{rows[-1]} ({len(rows)} rows)")


YELLOW = {"red": 1.0, "green": 1.0, "blue": 0.0}
WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}


def restyle_update_dates(svc, sid, sheet, today_iso):
    """Standing rule (2026-09-11): in the Update 날짜 / Exported Date column,
    today's cells are yellow-filled + bold; every other date is plain."""
    hdr = header_row(svc, sid, sheet)
    col = find_col(hdr, want_exact="Update 날짜") or find_col(hdr, want_exact="Exported Date")
    if not col:
        return
    props = sheet_props(svc, sid, sheet)
    letter = idx_to_a1_col(col)
    vals = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{sheet}'!{letter}2:{letter}",
        valueRenderOption="UNFORMATTED_VALUE").execute().get("values", [])
    y, m, d = (int(x) for x in today_iso.split("-"))
    serial = (datetime.date(y, m, d) - datetime.date(1899, 12, 30)).days
    kor = f"{y}. {m}. {d}"
    hits = [i + 2 for i, v in enumerate(vals)
            if v and (v[0] == serial or str(v[0]).strip() in (today_iso, kor))]
    last = max(len(vals) + 1, 2)
    gid = props["sheetId"]
    def fmt(r0, r1, bg, bold):
        return {"repeatCell": {
            "range": {"sheetId": gid, "startRowIndex": r0 - 1, "endRowIndex": r1,
                      "startColumnIndex": col - 1, "endColumnIndex": col},
            "cell": {"userEnteredFormat": {"backgroundColor": bg, "textFormat": {"bold": bold}}},
            "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold"}}
    reqs = [fmt(2, last, WHITE, False)]
    # collapse consecutive hit rows into ranges
    runs, start, prev = [], None, None
    for r in hits:
        if start is None: start = prev = r
        elif r == prev + 1: prev = r
        else: runs.append((start, prev)); start = prev = r
    if start is not None: runs.append((start, prev))
    reqs += [fmt(a, b, YELLOW, True) for a, b in runs]
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": reqs}).execute()
    print(f"[{sheet}] {letter}: cleared bold/fill on rows 2-{last}, yellow+bold on {len(hits)} "
          f"cells dated {today_iso}")


def phase_c_refresh_tem(svc, dry_run=True):
    """Rewrite tem cols F-K from each product's 1-5점 (유지훈P: 1-3점) Review-ID
    col K. Cols A-E are IMPORTRANGE — never touched. Safe to run every time."""
    cur = svc.spreadsheets().values().get(
        spreadsheetId=SRC, range=f"'{TEM_SHEET}'!A1:K"
    ).execute().get("values", [])
    cur_len = {}
    for ci in range(11):
        cur_len[idx_to_a1_col(ci + 1)] = sum(
            1 for row in cur[1:] if ci < len(row) and str(row[ci]).strip()
        )

    data = []
    for col, (sid, sheet, rid_col) in TEM_REFRESH.items():
        vals = svc.spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{sheet}'!{rid_col}2:{rid_col}"
        ).execute().get("values", [])
        ids = [v[0].strip() for v in vals if v and str(v[0]).strip() and v[0].strip() != "Review ID"]
        old_n = cur_len.get(col, 0)
        print(f"  tem!{col} ← {sheet}!{rid_col} : {old_n} → {len(ids)} ids")
        body = [[x] for x in ids] + [[""]] * max(0, TEM_CLEAR_TO_ROW - 1 - len(ids))
        data.append({"range": f"{TEM_SHEET}!{col}2:{col}{TEM_CLEAR_TO_ROW}", "values": body})

    if dry_run:
        print("  [dry-run] would rewrite tem cols F-K (A-E IMPORTRANGE untouched)")
        return
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=SRC,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()
    print("  ✓ tem cols F-K refreshed")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--new-sheet", help="SC_yymmdd or CaspiLM_yymmdd tab to funnel into SC (Phase A)")
    ap.add_argument("--product", choices=list(PRODUCTS), help="run Phase B/C/D for one product")
    ap.add_argument("--all-products", action="store_true", help="Phase B/D for all active products (dry-run only)")
    ap.add_argument("--refresh-tem", action="store_true", help="Phase C: rewrite tem cols F-K from the 1-5점/1-3점 Review-ID cols")
    ap.add_argument("--finish", action="store_true",
                    help="with --product: (re)run only the post-paste steps — =dr() on today's "
                         "un-classified rows + Update 날짜 restyle. Idempotent.")
    ap.add_argument("--commit", action="store_true", help="actually write (default: dry-run)")
    args = ap.parse_args()
    dry = not args.commit
    svc = get_service()

    if args.new_sheet:
        print("=== Phase A ===")
        phase_a_append_and_dedupe(svc, args.new_sheet, dry_run=dry)

    if args.refresh_tem:
        print("=== Phase C: tem refresh ===")
        phase_c_refresh_tem(svc, dry_run=dry)

    if args.all_products:
        prods = [p for p, c in PRODUCTS.items() if not c.get("inactive")]
    else:
        prods = [args.product] if args.product else []
    for p in prods:
        if args.finish:
            cfg = PRODUCTS[p]
            print(f"=== finish: {p} ===")
            if not cfg.get("dr_skip"):
                stamp_dr_on_filtered_13(svc, p, cfg, dry_run=dry)
            if not dry:
                restyle_update_dates(svc, cfg["dest_id"], cfg["dest_sheet"], today_kst_iso())
                if cfg.get("one_three_sheet"):
                    restyle_update_dates(svc, cfg["dest_id"], cfg["one_three_sheet"], today_kst_iso())
            else:
                print(f"[{p}] [dry-run] would restyle Update 날짜 on {cfg['dest_sheet']}"
                      + (f" + {cfg['one_three_sheet']}" if cfg.get("one_three_sheet") else ""))
            continue
        print(f"=== Phase B/C/D: {p} ===")
        phase_b_c_d_product(svc, p, dry_run=dry)

    if not args.new_sheet and not prods and not args.refresh_tem:
        ap.print_help()


if __name__ == "__main__":
    main()

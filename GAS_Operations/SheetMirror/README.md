# SheetMirror

Google Apps Script project that copies a large Google Sheet in chunks to a separate destination spreadsheet. Used to mirror `26년 전체문의` (the GCX master inquiry log, kept up to date by the `zendesk-inquiry-sync` job) to the `RAW` tab of a read-only dashboard sheet, with two cleaned helper columns re-applied as ARRAYFORMULAs.

**Script ID:** `1-t6Z95OM0EWsiXHOsExu-hPHMNRTHZ0HD7bdVoHMmbEPwKksRNyolVch`

## Screenshots

![`RAW` tab with the 🔄 Mirror menu](docs/raw_tab.jpg)
*`RAW` tab with the 🔄 Mirror menu*

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | `CONFIG`, `mirrorSheet()`, `onOpen()` menu, `colSegments_()`, `applyFormulas_()` |
| `appsscript.json` | GAS manifest (scope: `spreadsheets` only) |

---

## Config (top of `Code.js`)

| Key | Value |
|-----|-------|
| `sourceId` | `1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc` |
| `sourceSheet` | `26년 전체문의` |
| `destId` | `1qxwUjuV3-_0HRS1Bsb3Fsua0n8N6r6GzNnqiv9wRU10` |
| `destSheet` | `RAW` |
| `chunkSize` | `1000` rows per read-write cycle |
| `formulas` | map of column number → canonical ARRAYFORMULA, re-applied every run |

---

## Helper formula columns

The `RAW` sheet keeps two computed columns whose canonical formulas live in
`CONFIG.formulas` (keyed by **column number**). They are the **source of truth** —
`mirrorSheet()` skips them when clearing/writing mirrored data **and** re-applies
the formula to row 1 on every run, so they're always correct even if a cell was
edited or wiped:

- **Col 55 (BC1)** — `Brand(상세) Clean` → `={"Brand(상세) Clean";ARRAYFORMULA(REGEXREPLACE(E2:E, "Spigen\((.*?)\)", "$1"))}`
- **Col 56 (BD1)** — `Product Name Clean` → `={"Product Name Clean";ARRAYFORMULA(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(IF(ISBLANK(J2:J),K2:K,J2:J),"^\(.*?\)_","")," \(.*?\)",""),"_.*",""))}` — uses J (falls back to K), strips a leading `(…)_` prefix, any ` (…)` suffix and everything after `_`.

> Note: in `Code.js` the Brand formula's backslashes are doubled (JS string escaping); the cell receives single backslashes.
>
> ⚠️ Known issue: the col-56 formula is a JS **template literal** with single `\(` / `\)`, which JS turns into plain `(` / `)` — so the cell actually receives `"^(.*?)_"` and `" (.*?)"` (capture groups, not literal parentheses). `^(.*?)_` then strips everything up to the first `_`, and `" (.*?)"` only removes a single space. Double the backslashes (`\\(`) if the literal-parenthesis behaviour above is intended.

To change a formula or move it to another column, edit `CONFIG.formulas` in `Code.js` — not the cell — so the change survives the next mirror.

---

## Usage

Open the destination spreadsheet → **🔄 Mirror → Mirror now**, or run `mirrorSheet()` directly in the GAS editor.

How `mirrorSheet()` works:

1. `colSegments_()` splits columns `1..max(srcLastCol, dstLastCol)` into contiguous runs around the formula columns.
2. Clears only those non-formula segments in `RAW` (no whole-sheet `clearContents()`), so old trailing columns are cleared too.
3. Grows `RAW` rows/columns if the source is larger.
4. Reads the source in `chunkSize`-row chunks and writes each chunk per segment, skipping the formula columns. Data is mirrored by **absolute column position** — whatever the source has in cols 55/56 is not copied.
5. `applyFormulas_()` re-applies the helper formulas last so they spill over the freshly mirrored E / J / K data.

No trigger is defined in code; to automate, add a time-based trigger on `mirrorSheet` in the Apps Script UI.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_Operations/SheetMirror
clasp push --force
```

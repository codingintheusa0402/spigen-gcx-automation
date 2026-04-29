# Bi-Weekly Slide Updater

Google Apps Script project that auto-populates a bi-weekly CX report Google Slides deck
with live data from the `26년 전체문의` sheet.

## What it does

Each run replaces `{{placeholder}}` text boxes and inserts arc chart images directly
in the slide deck — no manual copy-paste needed.

### Placeholders replaced

| Placeholder | Value inserted |
|---|---|
| `{{TOTAL_INQUIRIES}}` | Total row count from `26년 전체문의` |
| `{{Defect_Reason_1}}` ~ `{{Defect_Reason_5}}` | Top 5 `인입사유` values under `Category = 4. Product Issue` |
| `{{Defect_Reason_1_Count}}` ~ `{{Defect_Reason_5_Count}}` | Corresponding counts |
| `{{Defect_Reason_<keyword>}}` | Count of rows whose `인입사유` contains `<keyword>` |
| `{{Defect_Model_Chart_1}}` ~ `{{Defect_Model_Chart_3}}` | Half-donut arc image (440×340 px) for top-3 defect products — **preserved on re-run** (see below) |
| `{{Defect_Model_Chart_Title_1}}` ~ `{{Defect_Model_Chart_Title_3}}` | Product name of top-N defect product |
| `{{Defect_Model_Chart_Count_1}}` ~ `{{Defect_Model_Chart_Count_3}}` | Total defect count with `건` suffix (e.g. `689건`) |
| `{{Defect_Model_Chart_Legend_1}}` ~ `{{Defect_Model_Chart_Legend_3}}` | 인입사유 names only, one per line (top 3 + 그 외) — put in a left-aligned text box |
| `{{Defect_Model_Chart_Legend_Value_1}}` ~ `{{Defect_Model_Chart_Legend_Value_3}}` | Corresponding counts, one per line — put in a right-aligned text box next to Legend |

### Chart placeholders (`{{Defect_Model_Chart_N}}`)

Place a text box containing exactly `{{Defect_Model_Chart_1}}` (or `_2`, `_3`) on a slide.
The script reads its position/size, clears its text, and inserts a 440×340 px PNG arc image
at the same position — sized to match the original text box.

**Preserve behavior**: once the placeholder text is cleared (i.e. the chart has been placed),
the script skips that slot on all subsequent runs — the image is preserved. To force a
refresh, retype `{{Defect_Model_Chart_N}}` into the (now empty) text box.

**Arc spec**
- Canvas: `440×340 px`
- Visible arc (half-circle): `220×110 px`, horizontally centered, `top = 45 px`
- Full circle chart area: `220×220`, `left = 110, top = 45`
- Invisible spacer half extends below the arc (y 155–265), same background color
- Shape: ∩ upward arch (`pieStartAngle: -90`, `circumference: 180°`)
- Rendered by built-in GAS `Charts` service — no external API calls
- Technique: spacer slice equal to visible data sum → forces data into exactly 180°; spacer colored `#11162d` (background)
- Colors: `#d336f4` / `#1554ff` / `#19c7f3` / `#8790b5` (reason 1–3 + 그 외)

### Legend placeholders

Use two side-by-side text boxes on the slide:

| Placeholder | Text box style | Example output |
|---|---|---|
| `{{Defect_Model_Chart_Legend_N}}` | Left-aligned | `황변`<br>`분리/이탈`<br>`자석탈락`<br>`그 외` |
| `{{Defect_Model_Chart_Legend_Value_N}}` | Right-aligned | `422`<br>`36`<br>`16`<br>`125` |

Both placeholders always produce the same number of lines so the two text boxes stay in sync.

## Source data

| Field | Value |
|---|---|
| Spreadsheet ID | `1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc` |
| Sheet name | `26년 전체문의` |
| Key columns | `Category`, `인입사유`, `Product Name` |
| Defect filter | `Category == "4. Product Issue"` |

## How to run

1. Open the linked Google Slides deck.
2. **Extensions → Apps Script** (or run from the GAS editor).
3. Run `updateSlideTextBoxes()`.
4. The slide deck is saved and closed automatically on completion.

Alternatively add a time-based trigger on `updateSlideTextBoxes` for fully automatic weekly updates.

## GAS Project

| Field | Value |
|---|---|
| Script ID | `1AXqBJHr-DITneMUV6BJcF6Zus5KtRudhQWRXh2e-Ei0nXhipqxJf4jVn` |
| Linked Slides | `14bH-E4YIhvHu13FiizHiDAl-nA4D52XpJU962raJA0E` |
| clasp push | `cd Bi-Weekly && clasp push --force` |
| clasp pull | `cd Bi-Weekly && clasp pull` |

## Functions

| Function | Purpose |
|---|---|
| `onOpen()` | Adds **Slide Updater → Update Slide Text** menu to the Slides UI |
| `updateSlideTextBoxes()` | Main entry point — orchestrates all text replacements and chart insertions |
| `buildTopProductsData(sheet, rowCount)` | Computes top-3 defect products with per-reason counts; result shared by text replacements and chart rendering |
| `buildLegendText(item)` | Returns reason names only, one per line (for `Legend_N` placeholder) |
| `buildLegendValues(item)` | Returns counts only, one per line (for `Legend_Value_N` placeholder) — mirrors `buildLegendText` line-for-line |
| `updateDefectModelCharts(presentation, topProducts)` | Iterates top-3 products and calls `insertChartAtPlaceholder` for each |
| `insertChartAtPlaceholder(...)` | If `{{}}` placeholder text box still exists: removes any prior auto-chart for that slot, inserts new PNG, clears placeholder text. If placeholder is already gone (chart preserved): no-op |
| `buildDefectModelChartBlob(data, title)` | Builds 440×340 half-donut arc PNG via GAS `Charts` service using the spacer-slice technique |
| `refreshLinkedCharts(presentation)` | Refreshes any Sheets-linked charts already embedded in the deck |
| `getColumnIndexByHeader(sheet, headerName)` | Looks up a column index by header name (1-based) |
| `extractKeywordPlaceholders(presentation, prefix)` | Scans all slides for `{{Defect_Reason_<keyword>}}` patterns |
| `findPlaceholderShapes(presentation, placeholder)` | Returns all Shape elements whose text contains the given placeholder string |
| `removeOldAutoCharts(presentation)` | Utility — removes all `AUTO_Defect_Model_Chart_*` images (not called automatically; use manually to wipe all charts at once) |

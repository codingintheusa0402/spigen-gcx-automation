# Bi-Weekly Slide Updater

Google Apps Script project that auto-populates a bi-weekly CX report Google Slides deck
with live data from the `26년 전체문의` sheet.

## What it does

Each run replaces `{{placeholder}}` text boxes and auto-inserts donut charts directly
in the slide deck — no manual copy-paste needed.

### Placeholders replaced

| Placeholder | Value inserted |
|---|---|
| `{{TOTAL_INQUIRIES}}` | Total row count from `26년 전체문의` |
| `{{Defect_Reason_1}}` ~ `{{Defect_Reason_5}}` | Top 5 `인입사유` values under `Category = 4. Product Issue` |
| `{{Defect_Reason_1_Count}}` ~ `{{Defect_Reason_5_Count}}` | Corresponding counts |
| `{{Defect_Reason_<keyword>}}` | Count of rows whose `인입사유` contains `<keyword>` |
| `{{Defect_Model_Chart_1}}` ~ `{{Defect_Model_Chart_3}}` | Donut chart image for top 3 defect products |

### Chart placeholders (`{{Defect_Model_Chart_N}}`)

Place a text box with exactly `{{Defect_Model_Chart_1}}` (or `_2`, `_3`) on a slide.
The script reads the text box's position/size, removes it, and inserts a donut chart
(PNG via Google Charts service) in its place — sized to match the original text box.

Charts show:
- Donut hole centre: product name + total defect count
- Up to 3 top `인입사유` slices + an "그 외" remainder slice
- Colors: `#d336f4` / `#1554ff` / `#19c7f3` / `#8790b5`

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
| `updateSlideTextBoxes()` | Main entry point — orchestrates all replacements |
| `updateDefectModelCharts()` | Builds per-product chart data and calls `insertChartAtPlaceholder` |
| `insertChartAtPlaceholder()` | Finds placeholder text box, swaps it for a PNG chart image |
| `buildDefectModelChartBlob()` | Builds donut chart PNG via `Charts` service |
| `refreshLinkedCharts()` | Refreshes any Sheets-linked charts already embedded in the deck |
| `getColumnIndexByHeader()` | Looks up a column index by header name (1-based) |
| `extractKeywordPlaceholders()` | Scans slides for `{{Defect_Reason_<keyword>}}` patterns |
| `findPlaceholderShapes()` | Returns all Shape elements containing a given placeholder string |
| `removeOldAutoCharts()` | Deletes previously inserted `AUTO_Defect_Model_Chart_*` images |
| `escapeSvg()` | HTML-escapes values for safe text insertion |

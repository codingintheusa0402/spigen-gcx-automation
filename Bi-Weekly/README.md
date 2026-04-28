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
| `{{Defect_Model_Chart_1}}` ~ `{{Defect_Model_Chart_3}}` | Half-donut arc image for top 3 defect products (arc only, no text) |
| `{{Defect_Model_Chart_Title_1}}` ~ `{{Defect_Model_Chart_Title_3}}` | Product name of top-N defect product |
| `{{Defect_Model_Chart_Count_1}}` ~ `{{Defect_Model_Chart_Count_3}}` | Total defect row count for that product |
| `{{Defect_Model_Chart_Legend_1}}` ~ `{{Defect_Model_Chart_Legend_3}}` | 4-line legend: top-3 `인입사유` + 그 외, tab-separated (see note below) |

### Chart placeholders (`{{Defect_Model_Chart_N}}`)

Place a text box with exactly `{{Defect_Model_Chart_1}}` (or `_2`, `_3`) on a slide.
The script reads the text box's position/size, removes it, and inserts a PNG arc image
in its place — sized to match the original text box.

The arc is a **∩ half-donut** (upward arch) rendered entirely by the built-in GAS
`Charts` service — no external API calls. Technique: a hidden spacer slice equal to
the visible data sum forces the real segments into exactly 180°; the spacer is
colored `#11162d` (background) so it disappears.

Arc colors: `#d336f4` / `#1554ff` / `#19c7f3` / `#8790b5` (reason 1–3 + 그 외)

### Legend placeholders (`{{Defect_Model_Chart_Legend_N}}`)

Produces 4 lines (3 top reasons + 그 외) in the format:

```
이유1\t75
이유2\t22
이유3\t13
그 외\t42
```

The `\t` tab character separates the reason name from the count. For the counts to
**right-align** at the edge of the text box, set a **right-aligned tab stop** at the
right edge of that text box in Google Slides:
*Select text box → Format → Bullets & numbering → set a right tab stop*.

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
| `buildTopProductsData(sheet, rowCount)` | Computes top-3 defect products with per-reason counts; shared by both text replacements and chart rendering |
| `buildLegendText(item)` | Formats a 4-line tab-separated legend string for one product |
| `updateDefectModelCharts(presentation, topProducts)` | Inserts arc chart images for the top-3 products |
| `insertChartAtPlaceholder(...)` | Finds a placeholder text box, removes it, inserts the PNG chart at the same position/size |
| `buildDefectModelChartBlob(data, title)` | Builds half-donut arc PNG via GAS `Charts` service (spacer-slice technique) |
| `refreshLinkedCharts(presentation)` | Refreshes any Sheets-linked charts already embedded in the deck |
| `getColumnIndexByHeader(sheet, headerName)` | Looks up a column index by header name (1-based) |
| `extractKeywordPlaceholders(presentation, prefix)` | Scans slides for `{{Defect_Reason_<keyword>}}` patterns |
| `findPlaceholderShapes(presentation, placeholder)` | Returns all Shape elements whose text contains the given placeholder |
| `removeOldAutoCharts(presentation)` | Deletes previously inserted `AUTO_Defect_Model_Chart_*` images before re-inserting |

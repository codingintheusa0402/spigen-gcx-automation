# Bi-Weekly Slide Updater

Google Apps Script project that auto-populates a bi-weekly CX report Google Slides deck
with live data from the `26년 전체문의` sheet.

**Scope**: all replacements and chart insertions run on the **currently active slide only**.
Switch to the target slide before running.

## What it does

Each run replaces `{{placeholder}}` text boxes and inserts arc chart images directly
in the slide deck — no manual copy-paste needed.

### Placeholders replaced

#### General

| Placeholder | Value inserted |
|---|---|
| `{{TOTAL_INQUIRIES}}` | Total row count from `26년 전체문의` |

#### Defect_Reason family (top defect reasons)

| Placeholder | Value inserted |
|---|---|
| `{{Defect_Reason_1}}` ~ `{{Defect_Reason_5}}` | Top 5 `인입사유` values under `Category = 4. Product Issue` |
| `{{Defect_Reason_1_Count}}` ~ `{{Defect_Reason_5_Count}}` | Corresponding counts |
| `{{Defect_Reason_<keyword>}}` | Count of rows whose `인입사유` contains `<keyword>` |

#### Defect_Model family — grouped by Product Name (top-3 products → top-3 reasons each)

| Placeholder | Value inserted |
|---|---|
| `{{Defect_Model_Chart_1}}` ~ `{{Defect_Model_Chart_3}}` | Half-donut arc image (440×340 px) for top-3 defect products — **preserved on re-run** (see below) |
| `{{Defect_Model_Chart_Title_1}}` ~ `{{Defect_Model_Chart_Title_3}}` | Product name of top-N defect product |
| `{{Defect_Model_Chart_Count_1}}` ~ `{{Defect_Model_Chart_Count_3}}` | Total defect count with `건` suffix (e.g. `689건`) |
| `{{Defect_Model_Chart_Legend_1}}` ~ `{{Defect_Model_Chart_Legend_3}}` | 인입사유 names only, one per line (top 3 + 그 외) |
| `{{Defect_Model_Chart_Legend_Value_1}}` ~ `{{Defect_Model_Chart_Legend_Value_3}}` | Corresponding counts, one per line |

#### Model_Defect family — grouped by 인입사유 (top-3 reasons → top-3 products each)

| Placeholder | Value inserted |
|---|---|
| `{{Model_Defect_Chart_1}}` ~ `{{Model_Defect_Chart_3}}` | Half-donut arc image (440×340 px) for top-3 defect reasons — **preserved on re-run** |
| `{{Model_Defect_Chart_Title_1}}` ~ `{{Model_Defect_Chart_Title_3}}` | 인입사유 name of top-N defect reason |
| `{{Model_Defect_Chart_Count_1}}` ~ `{{Model_Defect_Chart_Count_3}}` | Total count with `건` suffix |
| `{{Model_Defect_Chart_Legend_1}}` ~ `{{Model_Defect_Chart_Legend_3}}` | Product names only, one per line (top 3 + 그 외) |
| `{{Model_Defect_Chart_Legend_Value_1}}` ~ `{{Model_Defect_Chart_Legend_Value_3}}` | Corresponding counts, one per line |

#### Defect_Model_Glx26 family — same as Defect_Model, filtered to `Device` contains `'Galaxy S26'`

Covers Galaxy S26, S26+, S26 Ultra, etc. (substring match on the `Device` column).

| Placeholder | Value inserted |
|---|---|
| `{{Defect_Model_Chart_Glx26_1}}` ~ `{{Defect_Model_Chart_Glx26_3}}` | Half-donut arc image for top-3 defect products (Galaxy S26 rows only) — **preserved on re-run** |
| `{{Defect_Model_Chart_Title_Glx26_1}}` ~ `{{Defect_Model_Chart_Title_Glx26_3}}` | Product name |
| `{{Defect_Model_Chart_Count_Glx26_1}}` ~ `{{Defect_Model_Chart_Count_Glx26_3}}` | Total count with `건` suffix |
| `{{Defect_Model_Chart_Legend_Glx26_1}}` ~ `{{Defect_Model_Chart_Legend_Glx26_3}}` | 인입사유 names, one per line (top 3 + 그 외) |
| `{{Defect_Model_Chart_Legend_Value_Glx26_1}}` ~ `{{Defect_Model_Chart_Legend_Value_Glx26_3}}` | Corresponding counts, one per line |

#### Model_Defect_Glx26 family — same as Model_Defect, filtered to `Device` contains `'Galaxy S26'`

| Placeholder | Value inserted |
|---|---|
| `{{Model_Defect_Chart_Glx26_1}}` ~ `{{Model_Defect_Chart_Glx26_3}}` | Half-donut arc image for top-3 defect reasons (Galaxy S26 rows only) — **preserved on re-run** |
| `{{Model_Defect_Chart_Title_Glx26_1}}` ~ `{{Model_Defect_Chart_Title_Glx26_3}}` | 인입사유 name |
| `{{Model_Defect_Chart_Count_Glx26_1}}` ~ `{{Model_Defect_Chart_Count_Glx26_3}}` | Total count with `건` suffix |
| `{{Model_Defect_Chart_Legend_Glx26_1}}` ~ `{{Model_Defect_Chart_Legend_Glx26_3}}` | Product names, one per line (top 3 + 그 외) |
| `{{Model_Defect_Chart_Legend_Value_Glx26_1}}` ~ `{{Model_Defect_Chart_Legend_Value_Glx26_3}}` | Corresponding counts, one per line |

### Chart placeholders (`{{Defect_Model_Chart_N}}` / `{{Model_Defect_Chart_N}}`)

Place a text box containing exactly `{{Defect_Model_Chart_1}}` (or `_2`, `_3`, or the
`Model_Defect_` equivalents) on a slide.
The script reads its position/size, clears its text, and inserts a 440×340 px PNG arc image
at the same position — sized to match the original text box.

**Preserve behavior**: once the placeholder text is cleared (i.e. the chart has been placed),
the script skips that slot on all subsequent runs — the image is preserved. To force a
refresh, retype `{{Defect_Model_Chart_N}}` (or `{{Model_Defect_Chart_N}}`) into the
(now empty) text box.

**Arc spec**
- Canvas: `440×340 px`
- Visible arc (half-circle): `220×110 px`, horizontally centered, `top = 45 px`
- Full circle chart area: `220×220`, `left = 110, top = 45`
- Invisible spacer half extends below the arc (y 155–265), same background color
- Shape: ∩ upward arch (`pieStartAngle: -90`, `circumference: 180°`)
- Rendered by built-in GAS `Charts` service — no external API calls
- Technique: spacer slice equal to visible data sum → forces data into exactly 180°; spacer colored `#11162d` (background)
- Colors: `#d336f4` / `#1554ff` / `#19c7f3` / `#8790b5` (slot 1–3 + 그 외)

### Legend placeholders

Use two side-by-side text boxes on the slide:

**Defect_Model family** (names = 인입사유, values = counts):

| Placeholder | Text box style | Example output |
|---|---|---|
| `{{Defect_Model_Chart_Legend_N}}` | Left-aligned | `황변`<br>`분리/이탈`<br>`자석탈락`<br>`그 외` |
| `{{Defect_Model_Chart_Legend_Value_N}}` | Right-aligned | `422`<br>`36`<br>`16`<br>`125` |

**Model_Defect family** (names = Product Names, values = counts):

| Placeholder | Text box style | Example output |
|---|---|---|
| `{{Model_Defect_Chart_Legend_N}}` | Left-aligned | `Galaxy S26 Ultra`<br>`iPhone 17e`<br>`그 외` |
| `{{Model_Defect_Chart_Legend_Value_N}}` | Right-aligned | `312`<br>`98`<br>`44` |

Both Legend and Legend_Value placeholders always produce the same number of lines so the
two text boxes stay in sync.

## Source data

| Field | Value |
|---|---|
| Spreadsheet ID | `1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc` |
| Sheet name | `26년 전체문의` |
| Key columns | `Category`, `인입사유`, `Product Name`, `Device` |
| Defect filter | `Category == "4. Product Issue"` |
| Glx26 filter | `Device` contains `"Galaxy S26"` (applied on top of defect filter) |

## How to run

1. Open the linked Google Slides deck.
2. **Navigate to the slide you want to update.**
3. **Extensions → Apps Script** (or run from the GAS editor).
4. Run `updateSlideTextBoxes()`.
5. The slide deck is saved and closed automatically on completion.

All replacements, chart insertions, and linked-chart refreshes apply **only to the active
slide**. Run once per slide if you have multiple slides to update.

Alternatively add a time-based trigger on `updateSlideTextBoxes` for fully automatic weekly
updates (note: trigger runs may not have a user-selected slide context — test behavior first).

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
| `updateSlideTextBoxes()` | Main entry point — gets active slide, orchestrates all replacements and chart insertions on that slide only |
| `replaceTextOnSlide(slide, replacements)` | Iterates shapes on a single slide and applies all `{{key}} → value` substitutions in-place |
| `buildTopProductsData(sheet, rowCount, deviceFilter?)` | Computes top-3 defect products with per-reason counts; optional `deviceFilter` string restricts to rows whose `Device` col contains that text |
| `buildTopReasonsData(sheet, rowCount, deviceFilter?)` | Computes top-3 defect reasons with per-product counts; same optional `deviceFilter` |
| `buildLegendText(item)` | Returns reason names only, one per line (for `Defect_Model_Chart_Legend_N`) |
| `buildLegendValues(item)` | Returns counts only, one per line (for `Defect_Model_Chart_Legend_Value_N`) |
| `buildModelLegendText(item)` | Returns product names only, one per line (for `Model_Defect_Chart_Legend_N`) |
| `buildModelLegendValues(item)` | Returns counts only, one per line (for `Model_Defect_Chart_Legend_Value_N`) |
| `updateDefectModelCharts(slide, topProducts)` | Calls `insertChartAtPlaceholder` for each of top-3 products on the active slide |
| `updateModelDefectCharts(slide, topReasons)` | Calls `insertChartAtPlaceholder` for each of top-3 reasons on the active slide |
| `updateDefectModelChartsGlx26(slide, topProducts)` | Same as `updateDefectModelCharts` but uses `{{Defect_Model_Chart_Glx26_N}}` placeholders (Galaxy S26-filtered data) |
| `updateModelDefectChartsGlx26(slide, topReasons)` | Same as `updateModelDefectCharts` but uses `{{Model_Defect_Chart_Glx26_N}}` placeholders (Galaxy S26-filtered data) |
| `insertChartAtPlaceholder(slide, placeholder, chartData, title)` | If `{{}}` placeholder text box still exists on the slide: removes prior auto-chart for that slot, inserts new PNG, clears placeholder text. If placeholder is already gone (chart preserved): no-op |
| `buildDefectModelChartBlob(data, title)` | Builds 440×340 half-donut arc PNG via GAS `Charts` service using the spacer-slice technique |
| `refreshLinkedCharts(slide)` | Refreshes any Sheets-linked charts already embedded on the active slide |
| `findPlaceholderShape(slide, placeholder)` | Returns the first Shape on a slide whose text contains the given placeholder, or `null` |
| `findPlaceholderShapes(presentation, placeholder)` | Legacy — searches all slides; kept for manual use |
| `extractKeywordPlaceholders(slide, prefix)` | Scans active slide for `{{Defect_Reason_<keyword>}}` patterns |
| `getColumnIndexByHeader(sheet, headerName)` | Looks up a column index by header name (1-based) |
| `removeOldAutoCharts(presentation)` | Utility — removes all `AUTO_Defect_Model_Chart_*` images across all slides (not called automatically; use manually to wipe all charts at once) |

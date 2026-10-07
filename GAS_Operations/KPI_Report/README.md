# KPI_Report

Single-function, manual-run script that populates the team's half-year KPI narrative cells on the KPI tracking spreadsheet. Not a general-purpose reporting tool — the KPI text itself is hardcoded in the script and must be rewritten each period. A companion Python script, [`../generate_kpi_report.py`](../generate_kpi_report.py), produces the matching Word (.docx) KPI report.

**Script ID:** `1xu3wI3YRHy67BxC6E9-K10XIjy4BiGV9BoiUZksSmoxXAbmTPBiCki2X` (standalone)

---

## Files

| File | Purpose |
|------|---------|
| `populateKPIResults.gs` | `populateKPIResults()` — writes hardcoded KPI narrative strings into column L |
| `appsscript.json` | GAS manifest (scope: `spreadsheets` only) |
| `../generate_kpi_report.py` | Local python-docx generator for the H1 2026 KPI report document |

---

## What `populateKPIResults()` does

Opens a fixed spreadsheet/tab and writes 4 hardcoded Korean-language KPI result strings into `L4:L7` (row 3 is the header; order matches column C), each wrapped and top-aligned:

| Row | KPI (비중) |
|-----|-----------|
| L4 | KPI 1 — 아마존 리뷰 평점 개선 (20%) |
| L5 | KPI 2 — 제품 이슈 공론화 SIREN (30%) |
| L6 | KPI 3 — 고객경험·영업이익 개선 MCF (30%) |
| L7 | KPI 4 — CX 만족도 + 업무 자동화 (20%) |

| Key | Value |
|-----|-------|
| `SHEET_ID` | `15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI` |
| `SHEET_NAME` | `'26 상_김지우` |
| Target range | `L4:L7` |

No triggers — run manually from the Apps Script editor.

**To reuse for a new period:** edit the `kpiValues` array in `populateKPIResults.gs` with the new period's numbers/text, update `SHEET_NAME` if the tab changed (e.g. `'26 하_김지우`), then run `populateKPIResults()`.

---

## `generate_kpi_report.py` (Word report)

"Spigen GCX KPI Report Generator — Q1/Q2 2026 (v2)". Builds an A4 .docx with python-docx: per-KPI sections, before/after time tables, HR savings at the 2026 minimum wage (₩10,320/hr, `HOURLY`) plus platform/tool savings, and embedded screenshots.

```bash
pip install python-docx
python3 ~/Desktop/GCX/GAS_Operations/generate_kpi_report.py
# → ~/Desktop/GCX/Spigen_GCX_KPI_Report_2026_H1.docx
```

Screenshot inputs are absolute paths under `~/Desktop/` in the `IMG` dict — update them (and the numbers) per period before running.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_Operations/KPI_Report
clasp push --force
```

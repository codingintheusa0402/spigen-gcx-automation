# spigen-gcx-automation

Internal automation scripts for Spigen GCX (Global Customer Experience) — Amazon product data scraping and Google Sheets review monitoring.

## Project structure

```
spigen-gcx-automation/
├── APIFY_Axesso/               # GAS — Apify/Axesso review scraping & daily distribution
│   ├── Master.gs
│   ├── Code.gs
│   ├── Product.gs
│   └── README.md
├── amazon_dp_scraper/          # Python — Amazon /dp/ product detail scraper (Playwright)
│   ├── amazon_dp_scraper.py
│   └── README.md
└── amazon_child_asin_scraper/  # Python — Amazon child ASIN rating/review scraper (Selenium)
    ├── amazon_child_asin_scraper.py
    └── README.md
```

## Quick start

### Python scripts

```bash
pip install playwright openpyxl pynput selenium
playwright install chromium
```

### Google Apps Script

Deploy each `.gs` file into its corresponding Apps Script project. Set `APIFY_TOKEN` in **Project Settings → Script properties**.

## Branching conventions

| Branch | Use |
|--------|-----|
| `main` | Stable, production-ready |
| `feat/<desc>` | New features (e.g. `feat/add-ca-domain`) |
| `fix/<desc>` | Bug fixes (e.g. `fix/locale-gate-removed`) |

Commit message format: `<type>(<project>): <description>`
Examples: `fix(dp-scraper): remove locale gate`, `feat(apify-axesso): add drFormula field`

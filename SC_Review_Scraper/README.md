# Seller Central Review Scraper

Scrapes 1–3 star (critical) reviews from Amazon Seller Central US and enriches each review with customer-attached image URLs.

## How it works

1. **Page scraping** — Connects to an existing Chrome session via CDP and navigates through Seller Central review pages, extracting review content with human-like scroll simulation and random delays.
2. **Image enrichment** — Navigates to amazon.com and fetches each review's detail page using in-browser `fetch()` with session cookies to extract customer-attached image URLs.
3. **CSV output** — Writes all data to a UTF-8 BOM CSV (Excel-compatible).

## Prerequisites

### 1. Launch Chrome with remote debugging enabled

```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222
```

Or add `--remote-debugging-port=9222` to your Chrome shortcut/alias.

### 2. Log in to both accounts in that Chrome window

- **Amazon Seller Central** → `https://sellercentral.amazon.com/`
- **Google Chrome account** (sign in to your Chrome profile)

> Sessions expire overnight. If the script times out on page 1, re-login to both and re-run.

### 3. Install dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

## Configuration

Edit the constants at the top of `scrape_sc_reviews.py`:

| Variable | Default | Description |
|----------|---------|-------------|
| `PAGES` | `30` | Number of pages to scrape (50 reviews/page) |
| `OUT_FILE` | `~/Desktop/US_seller_central_reviews.csv` | Output CSV path |
| `NAV_DELAY` | `(2.0, 5.0)` | Random pause (s) between page loads |
| `BATCH_MIN/MAX` | `15/22` | Image-fetch batch size range |
| `FETCH_JITTER` | `(0, 600)` | Per-request stagger within a batch (ms) |

Full scrape of ~2,449 reviews: set `PAGES = 49`.

## Usage

```bash
python3 scrape_sc_reviews.py
```

## Output CSV fields

| Field | Description |
|-------|-------------|
| `ASIN` | Child ASIN of the reviewed product |
| `Created 날짜` | Review date |
| `사진 유무` | `Y` if customer attached images, `N` otherwise |
| `Reviewer` | Reviewer display name |
| `Review Ratings` | Star rating (1–5) |
| `Review Title` | Review headline |
| `본문` | Full review body text |
| `Product Rating` | Overall product star rating |
| `Ratings Count` | Total ratings count for the product |
| `Domain Code` | Marketplace (always `US` for this script) |
| `국가` | Country (always `US` for this script) |
| `Review Link` | Direct link to the review on amazon.com |
| `Image URL` | Pipe-delimited full-resolution image URLs (if any) |
| `Review ID` | Amazon review ID |

## Anti-bot measures

- Connects to an existing logged-in Chrome session (no bot browser fingerprint)
- Random delays between page navigations (2–5s)
- Human-like scroll simulation before each extraction
- Random batch sizes for image fetching (15–22 per batch)
- Per-request stagger within each batch (0–600ms jitter)
- Realistic browser headers on all amazon.com requests
- Same-origin `fetch()` with session cookies (indistinguishable from normal browsing)

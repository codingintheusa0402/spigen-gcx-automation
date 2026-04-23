# Seller Central Review Scraper

Scrapes reviews from Amazon Seller Central and enriches each review with customer-attached image URLs. Configurable per marketplace, star filter, detection avoidance level, and output columns.

## How it works

1. **Page scraping** — Connects to an existing Chrome session via CDP and navigates through Seller Central review pages, extracting review content with optional human-like scroll simulation and random delays.
2. **Deduplication** — Removes any duplicate Review IDs that appear across page boundaries before image fetching.
3. **Image enrichment** — Navigates to the marketplace's amazon domain and fetches each review's detail page using in-browser `fetch()` with session cookies to extract customer-attached image URLs.
4. **CSV output** — Writes all data to a UTF-8 BOM CSV (Excel-compatible, Korean-safe).

## Prerequisites

### 1. Launch Chrome with remote debugging enabled

```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222
```

### 2. Log in to both accounts in that Chrome window

- **Amazon Seller Central** for the target marketplace
- **Google Chrome account** (sign in to your Chrome profile)

> Sessions expire overnight. If the script times out on page 1, re-login to both and re-run.

### 3. Install dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

## Usage

```bash
python3 scrape_sc_reviews.py
```

---

## User Config

Edit the **USER CONFIG** section at the top of `scrape_sc_reviews.py`.

### `DOMAIN`

Marketplace to scrape.

| Value | Marketplace | Status |
|-------|-------------|--------|
| `"US"` | United States | Verified |
| `"UK"` | United Kingdom | Configured — verify SC URL before use |
| `"DE"` | Germany | Configured — verify SC URL before use |
| `"FR"` | France | Configured — verify SC URL before use |
| `"IT"` | Italy | Configured — verify SC URL before use |
| `"ES"` | Spain | Configured — verify SC URL before use |
| `"JP"` | Japan | Configured — verify SC URL before use |
| `"IN"` | India | Configured — verify SC URL before use |

> EU marketplaces (UK/DE/FR/IT/ES) share `sellercentral-europe.amazon.com`. Make sure the active Seller Central session is switched to the correct marketplace before running.

### `PAGES`

Number of pages to scrape. Each page returns up to 50 reviews.

```python
PAGES = 30      # ~1,500 reviews
PAGES = 49      # full US scrape (~2,449 reviews)
```

### `STAR_FILTER`

Comma-separated star ratings to include.

```python
STAR_FILTER = "1,2,3"       # critical reviews only (default)
STAR_FILTER = "1,2,3,4,5"   # all reviews
STAR_FILTER = "1"            # 1-star only
```

### `OUT_FILE`

Output CSV path. `None` auto-names the file based on domain.

```python
OUT_FILE = None                                      # → ~/Desktop/US_seller_central_reviews.csv
OUT_FILE = "/Users/me/Desktop/my_reviews.csv"        # custom path
```

### `STAR_FILTER`

Comma-separated star ratings to include.

```python
STAR_FILTER = "1,2,3"       # critical reviews only (default)
STAR_FILTER = "1,2,3,4,5"   # all reviews
STAR_FILTER = "1"            # 1-star only
```

### `ASIN_FILTER_FILE`

Path to a plain-text file with one ASIN per line. Only reviews matching those ASINs will appear in the output CSV. `None` saves all reviews regardless of ASIN.

```
# target_asins.txt
B0C9RT63VX
B096J9ZSG1
B0D52KPZ96
```

```python
ASIN_FILTER_FILE = None                                    # all ASINs (default)
ASIN_FILTER_FILE = "/Users/kevinkim/Desktop/target_asins.txt"
```

### `HEADERS_TO_INCLUDE`

Columns to keep in the output. `None` includes all 14 columns (default).

```python
HEADERS_TO_INCLUDE = None   # all columns

HEADERS_TO_INCLUDE = ['ASIN', 'Created 날짜', 'Reviewer', 'Review Ratings',
                       'Review Title', '본문', 'Image URL']
```

Full column list: `ASIN` · `Created 날짜` · `사진 유무` · `Reviewer` · `Review Ratings` · `Review Title` · `본문` · `Product Rating` · `Ratings Count` · `Domain Code` · `국가` · `Review Link` · `Image URL` · `Review ID`

### `DETECTION_AVOIDANCE`

Controls timing and behavioral patterns to avoid Amazon bot detection.

| Level | Nav delay | Batch delay | Jitter | Batch size | Scroll | Use when |
|-------|-----------|-------------|--------|------------|--------|----------|
| `"LOW"` | 0.5–1.5s | 0.5–1.5s | 0–150ms | 20–30 | No | Testing / one-off |
| `"MEDIUM"` | 2.0–5.0s | 2.0–4.5s | 0–600ms | 15–22 | Yes | Daily scheduled runs |
| `"HIGH"` | 4.0–10.0s | 5.0–12.0s | 0–1200ms | 8–15 | Yes | Large scrapes / high frequency |

---

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
| `Domain Code` | Marketplace code (e.g. `US`, `JP`) |
| `국가` | Country code |
| `Review Link` | Direct link to the review |
| `Image URL` | Pipe-delimited full-resolution image URLs (if any) |
| `Review ID` | Amazon review ID (used for deduplication) |

---

## Adding a new marketplace

Add an entry to `_DOMAINS` in the script:

```python
"CA": {
    "sc_base":     "https://sellercentral.amazon.ca/brand-customer-reviews/",
    "amazon_home": "https://www.amazon.ca/",
    "review_url":  "https://www.amazon.ca/gp/customer-reviews/",
    "country":     "CA",
},
```

Then set `DOMAIN = "CA"` and run.

---

## Anti-bot measures

- Connects to an existing logged-in Chrome session (no bot browser fingerprint)
- Randomized delays between page navigations
- Optional human-like scroll simulation before each extraction
- Random batch sizes for image fetching
- Per-request stagger (jitter) within each batch
- Realistic browser headers on all marketplace requests
- Same-origin `fetch()` with session cookies (indistinguishable from normal browsing)
- Deduplicates Review IDs before image fetching to prevent repeated requests

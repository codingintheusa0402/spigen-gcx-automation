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

| Value | Marketplace | Output file | Notes |
|-------|-------------|-------------|-------|
| `"US"` | United States | `US_*.csv` | Verified |
| `"EU"` | UK + DE + FR + IT + ES combined | `EU_*.csv` | Auto-expands; uses `marketplaceId` per country |
| `"UK"` | United Kingdom | `UK_*.csv` | Single-country run |
| `"DE"` | Germany | `DE_*.csv` | Single-country run |
| `"FR"` | France | `FR_*.csv` | Single-country run |
| `"IT"` | Italy | `IT_*.csv` | Single-country run |
| `"ES"` | Spain | `ES_*.csv` | Single-country run |
| `"JP"` | Japan | `JP_*.csv` | Verified |
| `"IN"` | India | `IN_*.csv` | Verified |

> `"EU"` automatically scrapes UK → DE → FR → IT → ES in sequence, switching marketplaces via the `marketplaceId` URL parameter. All five countries are written into one `EU_seller_central_reviews.csv`. Use individual codes (`"UK"`, `"DE"`, etc.) only if you need a single-country file.

### `PAGES`

Number of pages to scrape. Total reviews ≈ `PAGES × PAGE_SIZE`.

```python
PAGES = 30      # 1,500 reviews at PAGE_SIZE=50
PAGES = 25      # 2,500 reviews at PAGE_SIZE=100
```

### `PAGE_SIZE`

Number of reviews returned per page by Seller Central. Supported values: `25`, `50`, `100`.

```python
PAGE_SIZE = 50    # default — 50 reviews/page
PAGE_SIZE = 100   # fewer page loads, larger DOM per page
PAGE_SIZE = 25    # smaller batches
```

Higher values collect the same total reviews in fewer navigations but parse a larger DOM per page. `50` is the recommended default for stability.

### `STAR_FILTER`

Comma-separated star ratings to include.

```python
STAR_FILTER = "1,2,3"       # critical reviews only (default)
STAR_FILTER = "1,2,3,4,5"   # all reviews
STAR_FILTER = "1"            # 1-star only
```

### `OUT_DIR`

Directory where CSVs are saved. Each domain writes to `<OUT_DIR>/<DOMAIN>_seller_central_reviews.csv`.

```python
OUT_DIR = os.path.expanduser("~/Desktop")           # default
OUT_DIR = "/Users/me/Documents/reviews"             # custom path
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

### `HEADLESS` / `CHROME_USER_DATA`

Controls whether the browser is visible.

| Value | Behaviour | Use when |
|-------|-----------|----------|
| `False` (default) | Connects to your running Chrome via CDP (port 9222). Browser window is visible. | Debugging — watch the scraper navigate in real time |
| `True` | Launches a headless Chromium using your saved Chrome profile so existing login cookies are reused. No window appears. | Background / automated runs |

> **Headless requirement:** Chrome must be fully closed before running with `HEADLESS = True` — an open Chrome holds a lock on the profile directory that prevents Playwright from using it.

```python
HEADLESS = False
CHROME_USER_DATA = os.path.expanduser("~/Library/Application Support/Google/Chrome")
# Windows: "%LOCALAPPDATA%/Google/Chrome/User Data"
```

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

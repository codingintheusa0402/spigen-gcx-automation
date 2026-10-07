# Seller Central Review Scraper

Scrapes reviews from Amazon Seller Central (Brand Customer Reviews) and enriches each review with customer-attached image URLs and the Amazon Order ID. Top-level domains (US, EU, JP, IN) scrape in parallel; within EU, sub-countries scrape sequentially on one shared tab. At the end of a run all domain CSVs are merged into a dated `SC_{yymmdd}` tab of the source review spreadsheet, append-only and deduplicated by Review ID. Configurable per marketplace, star filter, detection avoidance level, and output columns.

The uploaded `SC_{yymmdd}` tab is then distributed to the product monitoring books by [`../SC_Master_Propagate`](../SC_Master_Propagate/README.md) (`sc-review-propagate` skill).

## Screenshots

![End of a real run (2026-10-07): per-marketplace review counts and the append-only `SC_261007` upload](docs/run_summary.jpg)
*End of a real run (2026-10-07): per-marketplace review counts and the append-only `SC_261007` upload*

## How it works

1. **Auto-launch Chrome** — Launches installed Google Chrome through Playwright's `launch_persistent_context(channel="chrome")` on a dedicated scraper profile (`~/.chrome-scraper-profile`). No CDP / remote-debugging port is used (CDP-attached sessions blocked downloads). Sessions persist between runs — log in once, done. On the unattended deployment (credentials file set) each account group gets its own profile, `~/.chrome-scraper-profile_<US|EU|JP|IN>`, because one Chrome profile can only hold one signed-in Amazon identity.
2. **Session check** — Navigates to each SC portal and checks if the session is still valid. Skips the login step entirely for portals that are already authenticated.
3. **Login tabs** — Only for portals that need login: opens one tab per SC endpoint (US, EU, JP, IN). Complete login + OTP on all tabs, then press Enter (interactive) or wait for the countdown (background run).
4. **Parallel scraping** — Top-level domains (US, EU, JP, IN) each get their own tab and scrape simultaneously. EU sub-countries (DE → IT → FR → ES → UK) run **sequentially** on one shared tab — all EU countries share the same SC Europe session cookie so parallel tabs would race each other. DE scrapes first; all remaining countries reuse the same tab, switching marketplace via the account-switcher dropdown before each country (the switcher only clicks "Spigen EU" when it is not already expanded, and dismisses the first-visit "Discover the new selling experience" modal — fix of 2026-10-07; before it, every DE/IT/FR/ES switch silently fell back to UK). Reusing one tab keeps the SC Europe session active throughout the entire EU run. If the session expires between countries anyway, the script detects the login redirect, pauses up to `MID_RUN_LOGIN_WAIT_SECONDS` (default 120 s) for you to complete OTP, then retries the marketplace switch and current page automatically — no data is lost.
5. **Incremental CSV write** — Reviews are flushed to CSV after every page so no data is lost if the run is interrupted.
6. **Deduplication** — Removes duplicate Review IDs across page boundaries before image fetching. Rows older than `MIN_REVIEW_DATE` and (optionally) ASINs outside `ASIN_FILTER_FILE` are dropped.
7. **Order ID** (`FETCH_ORDER_ID`, default on) — calls Seller Central's internal `brandcustomerreviews/api/reviews` endpoint with the existing SC cookies to add the Amazon Order ID of verified-purchase reviews.
8. **Image enrichment** — Navigates to the Amazon domain and fetches each review's detail page using in-browser `fetch()` with session cookies to extract customer-attached image URLs. **EU limitation**: only DE reviews get image URLs because the scraper Chrome profile has a customer session on amazon.de only. IT, FR, ES, and UK are skipped for image fetch until customer sessions for those domains are added to the profile.
9. **Google Sheets upload** (`UPLOAD_TO_SHEETS`, default on) — combines all domain CSVs (EU → JP → US → IN) and writes them to the `SC_{yymmdd}` tab (KST date, or `SC_SCRAPER_RUN_DATE`) of spreadsheet `SHEETS_SPREADSHEET_ID`. If the tab doesn't exist yet it is created; if it does (same-day re-run) only rows whose Review ID isn't already in it are appended — existing rows are never rewritten. Prints `[XX] new N | already in sheet M` per country. Uses the OAuth token at `~/.config/gws_shim/token.json`.

## Prerequisites

```bash
pip install -r requirements.txt
playwright install chromium
```

> Chrome must be installed at `/Applications/Google Chrome.app` (default Mac path). Update `CHROME_PATH` in the config if yours differs.

## Usage

```bash
python3 scrape_sc_reviews.py

# smoke test: one domain, few pages, no image/order-ID fetch, no Sheets write
SC_SCRAPER_DOMAINS=US SC_SCRAPER_PAGES=1 SC_SCRAPER_FETCH_IMAGES=0 \
SC_SCRAPER_FETCH_ORDER_ID=0 SC_SCRAPER_UPLOAD=0 python3 scrape_sc_reviews.py
```

There are no CLI flags — everything is the USER CONFIG block plus the env overrides below. **Set `SC_SCRAPER_UPLOAD=0` for any test run**; the default writes a live `SC_{yymmdd}` tab to the production spreadsheet.

Or use the `/sc-scraper` Claude Code skill — it asks for all options interactively, edits the config, and runs the script automatically.

On first run Chrome opens automatically → log in to all SC accounts → press Enter. Subsequent runs reuse saved sessions and start scraping immediately. When started in the background the log goes to `/tmp/sc_scraper.log`.

### Environment overrides

| Env var | Default | Effect |
|---|---|---|
| `SC_SCRAPER_DOMAINS` | `EU,JP,US,IN` | Comma-separated domain list |
| `SC_SCRAPER_PAGES` | `30` | Page limit per domain |
| `SC_SCRAPER_FETCH_IMAGES` | `1` | `0` skips image enrichment |
| `SC_SCRAPER_FETCH_ORDER_ID` | `1` | `0` skips the Order ID lookup |
| `SC_SCRAPER_UPLOAD` | `1` | `0` skips the Google Sheets upload |
| `SC_SCRAPER_RUN_DATE` | KST today | `yymmdd` used for the `SC_<date>` tab name (backfills / tests) |
| `SC_SCRAPER_HEADLESS` | `0` | `1` = headless Chrome (servers with no display) |
| `SC_SCRAPER_OUT_DIR` | `~/Desktop` | CSV output dir |
| `SC_SCRAPER_SCREENSHOT_DIR` | `~/sc_scraper_screenshots` | Login-failure / diagnostic screenshots |
| `SC_SCRAPER_CREDENTIALS_FILE` | unset | Enables automated login (see below). Never set on the Mac |
| `SC_SCRAPER_CHAT_WEBHOOK` | unset | Private Chat space for OTP requests + failure alerts (read by `sc_auth.py`) |
| `SC_SCRAPER_DIAGNOSE_ACCOUNTS` / `SC_SCRAPER_ISOLATED_TEST_DOMAIN` / `SC_SCRAPER_DIAGNOSE_CUSTOMER_LOGIN` | unset | Diagnostic-only modes: dump the account-picker HTML, try a fresh throwaway-profile login for one domain, or test only the storefront customer login for one domain — then exit without scraping |

---

## Hourly scheduling (currently disabled)

`run_hourly.sh` is the wrapper for the `com.spigen.sc-scraper.hourly` LaunchAgent (`StartInterval` 3600). It skips a cycle if a previous run is still going, unsets the server-only env vars, runs the scraper with `/opt/homebrew/bin/python3`, and appends the per-country `new N | already in sheet M` lines to a permanent history log.

| Log | Content |
|---|---|
| `/tmp/sc_scraper.log` | Full output of the latest run (overwritten each cycle) |
| `/tmp/sc_scraper_hourly.log` | Trigger / skip / exit-code lines |
| `~/.sc_scraper_new_reviews_history.log` | Per-run new-review counts, kept indefinitely |

The plist is not loaded at the moment — it is parked in the local, untracked `disabled_launchagents/` folder. To re-enable, copy it to `~/Library/LaunchAgents/` and `launchctl bootstrap gui/$(id -u) <plist>`.

---

## Unattended deployment (EC2)

For running daily without a local machine, set the `SC_SCRAPER_CREDENTIALS_FILE` env var to a local file (not committed anywhere) with one `DOMAIN|EMAIL|PASSWORD` line per top-level domain group (`US`, `EU`, `JP`, `IN` — EU covers all 5 sub-countries via the shared SC Europe login):

```
US|<us-seller-central-email>|<password>
EU|<eu-seller-central-email>|<password>
JP|<jp-seller-central-email>|<password>
IN|<in-seller-central-email>|<password>
```

When set, `sc_auth.py` automatically fills email + password on any login/re-login prompt. It deliberately does **not** auto-generate the OTP from a stored TOTP secret — that would give a compromised server permanent, silent MFA bypass. Instead, whenever Amazon asks for a one-time code, it opens a `cloudflared` quick tunnel, posts a link to a private Google Chat webhook, and waits for a human to submit the live code from Google Authenticator via a small one-time web form. On the Mac, this env var is never set, so login stays fully manual as described above.

Other env vars honored on the unattended deployment: `SC_SCRAPER_OUT_DIR` (CSV output dir, no `~/Desktop` on a server), `SC_SCRAPER_SCREENSHOT_DIR` (login-failure screenshots), `SC_SCRAPER_CHAT_WEBHOOK` (OTP-request + failure-alert messages — use a private space, not a shared team room), `SC_SCRAPER_HEADLESS=1`.

### Apify Actor packaging

`.actor/` (`actor.json`, `input_schema.json`, `Dockerfile`) + `main.py` package the same scraper as the Apify Actor `sc-review-scraper`. `main.py` is a thin wrapper: it turns the Actor's secret inputs (`US/EU/JP/IN_EMAIL` + `_PASSWORD`, `CHAT_WEBHOOK_URL`, `GWS_TOKEN_JSON`, diagnostic toggles) into the env vars / credentials file above, restores each domain's Chrome profile from the Key-Value Store `sc-scraper-state` before the run and saves it back afterwards (cache dirs trimmed), and pushes the uploaded rows to the Actor's default dataset so they can be exported from the run's Output tab. `scrape_sc_reviews.py` itself runs unmodified.

### 24/7 server

The same script runs on the GCX Windows/WSL2 server (see [`../../ServerBootstrap`](../../ServerBootstrap/README.md)); `ServerBootstrap/test/sc_test_report.py` waits for a server run to finish and posts a result card to the private test room only.

---

## User Config

Edit the **USER CONFIG** section at the top of `scrape_sc_reviews.py`.

### `DOMAINS`

Marketplaces to scrape in parallel.

| Value | Marketplace | Output file | Notes |
|-------|-------------|-------------|-------|
| `"US"` | United States | `US_*.csv` | |
| `"EU"` | UK + DE + FR + IT + ES combined | `EU_*.csv` | Auto-expands; sub-countries scrape sequentially on one tab |
| `"UK"` | United Kingdom | `UK_*.csv` | Single-country run |
| `"DE"` | Germany | `DE_*.csv` | Single-country run |
| `"FR"` | France | `FR_*.csv` | Single-country run |
| `"IT"` | Italy | `IT_*.csv` | Single-country run |
| `"ES"` | Spain | `ES_*.csv` | Single-country run |
| `"JP"` | Japan | `JP_*.csv` | |
| `"IN"` | India | `IN_*.csv` | |

Default (all markets): `DOMAINS = ["EU", "JP", "US", "IN"]` (or `SC_SCRAPER_DOMAINS`). Use `"EU"`, not a bare `"DE"`, for a normal run — a bare sub-country is treated as its own login group.

### `EU_COUNTRIES`

Controls which EU sub-countries are scraped when `"EU"` is in `DOMAINS`. Defaults to all five.

```python
EU_COUNTRIES = ["DE", "IT", "FR", "ES", "UK"]   # all (default)
EU_COUNTRIES = ["IT"]                             # Italy-only re-run
```

DE is always scraped first (Phase 1) if it's in the list. The remaining countries follow sequentially on the same tab (Phase 2). Removing DE from the list skips Phase 1 entirely — useful for appending a missed country to an existing EU CSV.

**Single-country re-run example** (append Italy to an existing EU CSV):
```python
DOMAINS       = ["EU"]
EU_COUNTRIES  = ["IT"]
APPEND_CSV    = True
PAGES         = 50
```

### `PAGES` / `PAGES_OVERRIDE`

`PAGES` is the default page limit per domain. `PAGES_OVERRIDE` lets you set different limits per domain.

```python
PAGES = 30                                    # default (SC_SCRAPER_PAGES) — 1,500 reviews at PAGE_SIZE=50
PAGES_OVERRIDE = {}                           # no overrides (default)
PAGES_OVERRIDE = {"US": 49, "JP": 10}        # US gets 49 pages, JP gets 10, others use PAGES
```

### `PAGE_SIZE`

Reviews per page. Supported: `25`, `50`, `100`.

```python
PAGE_SIZE = 50    # default — 50 reviews/page
```

### `START_PAGE` / `APPEND_CSV`

Resume an interrupted run without losing already-saved rows.

```python
START_PAGE = 1        # start from the beginning (default)
APPEND_CSV = False    # overwrite CSV on start (default)

# Resume example — pick up from page 20, keep existing rows:
START_PAGE = 20
APPEND_CSV = True
```

`APPEND_CSV = True` also controls EU Phase 1: when DE is in `EU_COUNTRIES`, setting `APPEND_CSV = True` appends DE rows to an existing EU CSV instead of rewriting it.

### `STAR_FILTER`

```python
STAR_FILTER = "1,2,3,4,5"   # all reviews (default)
STAR_FILTER = "1,2,3"        # critical reviews only
```

### `MIN_REVIEW_DATE`

`yyyy-mm-dd` (currently `"2026-07-25"`) or `None`. Only reviews created on/after this date are kept.

### `FETCH_IMAGES` / `FETCH_ORDER_ID`

```python
FETCH_IMAGES = True    # fetch reviewer-attached image URLs (default; SC_SCRAPER_FETCH_IMAGES=0 to skip)
FETCH_ORDER_ID = True  # add the Order ID column via the SC internal API (default; SC_SCRAPER_FETCH_ORDER_ID=0 to skip)
```

### `UPLOAD_TO_SHEETS` / `SHEETS_SPREADSHEET_ID`

`True` by default (`SC_SCRAPER_UPLOAD=0` to skip). Target spreadsheet `1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw`, tab `SC_{yymmdd}`; append-only dedupe by Review ID as described above. Token: `~/.config/gws_shim/token.json`.

### `HEADLESS`

`False` on the Mac (visible Chrome). `SC_SCRAPER_HEADLESS=1` for servers; Chrome must be fully closed first.

### `FETCH_IMAGES_ONLY`

Crash-recovery mode. Set to `True` to skip all scraping and re-run only the image fetch phase on already-saved CSVs.

```python
FETCH_IMAGES_ONLY = False   # normal run (default)
FETCH_IMAGES_ONLY = True    # skip scraping, re-fetch images on existing CSVs
```

### `MID_RUN_LOGIN_WAIT_SECONDS`

Seconds to wait when a login redirect is detected mid-scrape (e.g. session expired between EU countries). The script pauses, prints a warning, and retries automatically after the timer — no pages are skipped. Default: `120`.

### `DETECTION_AVOIDANCE`

Default: `"LOW"`.

| Level | Nav delay | Batch delay | Jitter | Batch size | Scroll | Use when |
|-------|-----------|-------------|--------|------------|--------|----------|
| `"LOW"` | 0.5–1.5s | 0.5–1.5s | 0–150ms | 20–30 | No | Testing / one-off |
| `"MEDIUM"` | 2.0–5.0s | 2.0–4.5s | 0–600ms | 15–22 | Yes | Daily scheduled runs |
| `"HIGH"` | 4.0–10.0s | 5.0–12.0s | 0–1200ms | 8–15 | Yes | Large scrapes / high frequency |

### `ASIN_FILTER_FILE`

Path to a plain-text file with one ASIN per line. Only reviews matching those ASINs appear in the output. `None` (default) saves all reviews. Local ASIN lists are kept in `filters/` (untracked).

### `OUT_DIR`

Directory where CSVs are saved (`<OUT_DIR>/<DOMAIN>_seller_central_reviews.csv`). Default: `~/Desktop`, or `SC_SCRAPER_OUT_DIR`.

### `HEADERS_TO_INCLUDE`

Columns to keep in the output, in this exact order. `None` includes all 15 columns. Default (only `Domain Code` excluded; this order is the 14-column A..N layout the `SC` master sheet and `SC_Master_Propagate` expect):

```python
HEADERS_TO_INCLUDE = [
    'ASIN', 'Created 날짜', '사진 유무', 'Reviewer', 'Review Ratings',
    'Review Title', '본문', '국가', 'Review Link', 'Image URL', 'Review ID',
    'Order ID', 'Product Rating', 'Ratings Count',
]
```

Full column list: `ASIN` · `Created 날짜` · `사진 유무` · `Reviewer` · `Review Ratings` · `Review Title` · `본문` · `Product Rating` · `Ratings Count` · `Domain Code` · `국가` · `Review Link` · `Image URL` · `Review ID` · `Order ID`

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
| `Order ID` | Amazon order ID for verified purchases (empty otherwise) |

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

Then add `"CA"` to `DOMAINS` and run.

---

## Anti-bot measures

- Real installed Chrome on a persistent, logged-in profile (no fresh-browser fingerprint)
- Randomized delays between page navigations
- Human-like scroll simulation before each extraction (MEDIUM / HIGH)
- Random batch sizes for image fetching
- Per-request stagger (jitter) within each batch
- Realistic browser headers on all marketplace requests
- Same-origin `fetch()` with session cookies (indistinguishable from normal browsing)
- Deduplicates Review IDs before image fetching to prevent repeated requests

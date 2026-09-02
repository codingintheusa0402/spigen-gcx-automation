#!/usr/bin/env python3
"""Scrape reviews from Amazon Seller Central with configurable options.

Edit the USER CONFIG section below, then run:
    python3 scrape_sc_reviews.py
"""

import asyncio, csv, random, os, sys, json, time, shutil
from collections import defaultdict
from datetime import datetime, timezone, timedelta
sys.stdout.reconfigure(line_buffering=True)  # flush every print immediately when running in background
from playwright.async_api import async_playwright
import requests
from sc_auth import load_credentials, ensure_logged_in

# Unattended-login support (EC2/server deployment only). Unset on the Mac —
# _creds stays empty there, so ensure_logged_in() is never called and the
# original manual input()/sleep-loop login flow is completely unchanged.
CREDENTIALS_FILE = os.environ.get("SC_SCRAPER_CREDENTIALS_FILE")
_creds = load_credentials(CREDENTIALS_FILE) if CREDENTIALS_FILE else {}
SCREENSHOT_DIR = os.environ.get("SC_SCRAPER_SCREENSHOT_DIR", os.path.expanduser("~/sc_scraper_screenshots"))
DIAGNOSE_ACCOUNTS = os.environ.get("SC_SCRAPER_DIAGNOSE_ACCOUNTS", "0") == "1"
ISOLATED_TEST_DOMAIN = os.environ.get("SC_SCRAPER_ISOLATED_TEST_DOMAIN", "")
DIAGNOSE_CUSTOMER_LOGIN = os.environ.get("SC_SCRAPER_DIAGNOSE_CUSTOMER_LOGIN", "")
# When set to a domain code (US/DE/JP/IN — a top-level domain, not "EU"),
# skips Seller Central and the full scrape entirely: opens that domain's
# REAL persisted per-domain profile (the one production image-fetch runs
# against, not a throwaway one), navigates straight to the customer
# storefront, runs ensure_customer_logged_in(), and dumps a screenshot +
# HTML of wherever it lands, before and after. Used to see exactly what
# Amazon serves when the storefront login flow fails, without paying for a
# full scrape run first.
# When set to a domain code (US/EU/JP/IN), forces a completely fresh,
# throwaway Chrome profile (never the shared persistent one) and attempts a
# genuine login using only that domain's credentials — used to test whether
# a domain's session-check "already logged in" result on the shared profile
# is a real login or just inherited cross-domain session cookies from a
# different account. Never set on the Mac.
# Diagnostic-only mode: dump full page HTML (and any "Switch Accounts"-style
# dropdown, expanded) for each domain's landing page immediately after
# login/session-check, then exit — skips the entire scrape/upload pipeline.
# Used to inspect Amazon's account-picker structure without burning API
# calls or scrape time against real accounts. Never set on the Mac.

LAST_COMBINED_ROWS = None
# Set by main() at the end of a run to the same header+rows list uploaded to
# Google Sheets (or None if UPLOAD_TO_SHEETS is off, or upload never got far
# enough to build it). main.py (Apify deployment only) reads this after
# calling main() — including after a caught SystemExit from a partial-domain
# failure — to push the identical data to the Actor's default dataset, so
# it's downloadable from the Runs/Output tab too. Unused on the Mac.

# ═══════════════════════════════════════════════════════════════════════════════
# USER CONFIG — edit these before each run
# ═══════════════════════════════════════════════════════════════════════════════

DOMAINS = ["EU", "JP", "US", "IN"]
# List of domains to scrape in parallel. Each gets its own CSV file.
# Single domain example : DOMAINS = ["US"]
# Supported             : "US" | "EU" | "UK" | "DE" | "FR" | "IT" | "ES" | "JP" | "IN"
# "EU" automatically scrapes UK + DE + FR + IT + ES in sequence using each
# country's marketplaceId and writes all reviews into one EU_*.csv file.

PAGES = 5
# Default max pages to scrape per domain.
# Total reviews ≈ PAGES × PAGE_SIZE.
# Override per-domain with PAGES_OVERRIDE below.

PAGES_OVERRIDE = {}
# Per-domain page limit. Domains not listed here use PAGES.
# UK and DE are EU sub-countries — their overrides apply when scraping "EU" too.

PAGE_SIZE = 50
# Number of reviews per page returned by Seller Central.
# Supported values: 25 | 50 | 100
# Higher values mean fewer page loads but larger DOM per page.

START_PAGE = 1
# Page to start from. Set > 1 to resume a previously interrupted run.
# When resuming, also set APPEND_CSV = True to avoid overwriting saved rows.

APPEND_CSV = False
# False — overwrites the CSV at the start (default, fresh run).
# True  — appends to an existing CSV without rewriting the header.
#          Use together with START_PAGE to resume an interrupted run.

STAR_FILTER = "1,2,3,4,5"
# Comma-separated star ratings to include.
# Critical reviews only: "1,2,3"   All reviews: "1,2,3,4,5"

OUT_DIR = os.environ.get("SC_SCRAPER_OUT_DIR", os.path.expanduser("~/Desktop"))
# Directory where CSVs are saved.
# Each domain is saved as <OUT_DIR>/<DOMAIN>_seller_central_reviews.csv
# Overridable via SC_SCRAPER_OUT_DIR env var (used on the EC2 deployment,
# which has no ~/Desktop) — defaults to ~/Desktop unchanged on the Mac.

HEADERS_TO_INCLUDE = [
    'ASIN', 'Created 날짜', '사진 유무', 'Reviewer', 'Review Ratings',
    'Review Title', '본문', '국가', 'Review Link', 'Image URL', 'Review ID',
    'Order ID', 'Product Rating', 'Ratings Count',
]
# Columns to keep in the output CSV, and in this exact order. None → all columns.
# Default (set 2026-08-18): 'Product Rating'/'Ratings Count' (the ASIN's overall
# rating + total rating count, shown as "Rating"/"Review Count" in the sheet)
# placed after 'Order ID' (cols M/N) per standing user preference.
# 'Domain Code' remains excluded per the earlier 2026-08-06 preference.
# Full list: ASIN | Created 날짜 | 사진 유무 | Reviewer | Review Ratings | Review Title
#            본문 | Product Rating | Ratings Count | Domain Code | 국가
#            Review Link | Image URL | Review ID | Order ID

ASIN_FILTER_FILE = None
# Path to a .txt file containing one ASIN per line.
# Only reviews whose ASIN matches an entry in this file will be saved to CSV.
# None → save all reviews regardless of ASIN (default).
# Example: ASIN_FILTER_FILE = "/Users/kevinkim/Desktop/target_asins.txt"

MIN_REVIEW_DATE = "2026-07-25"
# yyyy-mm-dd string, or None to disable. Only reviews with Created 날짜 >= this
# date are kept (applied after ASIN filtering, before CSV write).

FETCH_IMAGES = True
# True  — fetch reviewer-attached media URLs after all page scraping is done.
#         EU fetches images for all countries in one pass after UK→DE→FR→IT→ES.
# False — skip image fetching entirely (faster runs, Image URL column stays empty).

FETCH_IMAGES_ONLY = False
# True  — skip all scraping; just re-run the image fetch on existing CSV files.
#         Use this to recover after a browser crash that interrupted the image fetch.
#         DOMAINS and OUT_DIR must match the original run so the right CSVs are found.

FETCH_ORDER_ID = True
# True  — after scraping each domain, call the Seller Central internal API
#         (brandcustomerreviews/api/reviews) to retrieve the Amazon Order ID for
#         each verified-purchase review. Adds an 'Order ID' column to the CSV.
#         ~84 % of DE reviews have an Order ID; non-verified purchases return "".
#         Reuses the same browser context as scraping. No extra login needed —
#         reuses the existing SC cookies.
# False — skip order ID fetching (default, faster runs).

UPLOAD_TO_SHEETS = True
# True  (default) — after scraping, combine all domain CSVs (EU+JP+US+IN, in that
#         scrape order) and merge them into the SC_{yymmdd} worksheet (KST date
#         of the run) on SHEETS_SPREADSHEET_ID. If that sheet doesn't exist yet
#         today, it's created fresh. If it already exists (e.g. an hourly re-run),
#         only rows whose Review ID isn't already in the sheet are appended —
#         existing rows are never touched, rewritten, or duplicated. Per-country
#         (국가) new-vs-already-present counts are printed each run.
# False — skip upload.

SHEETS_SPREADSHEET_ID = "1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw"
# Target Google Spreadsheet for UPLOAD_TO_SHEETS. Credentials: ~/.config/gws_shim/token.json.

LOGIN_WAIT_SECONDS = 300
# Seconds to wait for manual login when running non-interactively (background / no TTY).
# In interactive mode the script waits for Enter instead (no fixed timeout).

MID_RUN_LOGIN_WAIT_SECONDS = 120
# Seconds to wait when a login redirect is detected mid-scrape (session expired).
# The script pauses, lets you complete OTP, then retries the same page.
# In interactive mode it waits for Enter instead.

DETECTION_AVOIDANCE = "LOW"
# LOW    — short delays, fastest runs, higher detection risk
# MEDIUM — randomized delays + scroll simulation (recommended for daily use)
# HIGH   — aggressive randomization + long delays (safest for large/frequent scrapes)

HEADLESS = os.environ.get("SC_SCRAPER_HEADLESS", "0") == "1"
# False (default) — auto-launches Chrome with SCRAPER_PROFILE_DIR; sessions persist
#                   between runs so you only need to log in once. Browser is visible.
# True            — launches headless Chromium using SCRAPER_PROFILE_DIR.
# Overridable via SC_SCRAPER_HEADLESS=1 env var (set on unattended cloud
# deployments, which have no display) — defaults to visible/False on the Mac.
#                   Chrome must be fully closed before running in headless mode.

SCRAPER_PROFILE_DIR = os.path.expanduser("~/.chrome-scraper-profile")
# Dedicated Chrome profile for scraping. SC login sessions are saved here between runs.
# First run: Chrome opens → log in to all SC accounts → sessions persist automatically.
# Subsequent runs: Chrome opens with saved sessions → scraping starts after Enter.

# ═══════════════════════════════════════════════════════════════════════════════
# DOMAIN REGISTRY
# Add new marketplaces here as they are onboarded.
# EU domains (UK/DE/FR/IT/ES) share sellercentral-europe.amazon.com but each
# requires its own marketplaceId parameter. When DOMAINS = ["EU"], the scraper
# automatically loops through all EU_COUNTRIES and writes to one EU_*.csv.
# ═══════════════════════════════════════════════════════════════════════════════

# When "EU" is in DOMAINS, these sub-countries are scraped in order.
EU_COUNTRIES = ["DE", "IT", "FR", "ES", "UK"]
# DE is scraped first (alone); IT, FR, ES, UK then follow sequentially.

_DOMAINS = {
    "US": {
        "sc_base":     "https://sellercentral.amazon.com/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.com/",
        "review_url":  "https://www.amazon.com/gp/customer-reviews/",
        "country":     "US",
    },
    # EU countries all share sellercentral-europe.amazon.com (one login session).
    # sc_display_name: account-switcher dropdown label for this country.
    # sc_mkid: Amazon standard marketplace ID — used to construct mons_sel_* URL params
    #          directly without relying on the dropdown UI for every country.
    "UK": {
        "sc_base":         "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "sc_display_name": "United Kingdom",
        "sc_mkid":         "A1F83G8C2ARO7P",
        "amazon_home":     "https://www.amazon.co.uk/",
        "review_url":      "https://www.amazon.co.uk/gp/customer-reviews/",
        "country":         "UK",
        "image_fetch":     False,  # no amazon.co.uk customer session in scraper profile
    },
    "DE": {
        "sc_base":         "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "sc_display_name": "Germany",
        "sc_mkid":         "A1PA6795UKMFR9",
        "amazon_home":     "https://www.amazon.de/",
        "review_url":      "https://www.amazon.de/gp/customer-reviews/",
        "country":         "DE",
    },
    "FR": {
        "sc_base":         "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "sc_display_name": "France",
        "sc_mkid":         "A13V1IB3VIYZZH",
        "amazon_home":     "https://www.amazon.fr/",
        "review_url":      "https://www.amazon.fr/gp/customer-reviews/",
        "country":         "FR",
        "image_fetch":     False,  # no amazon.fr customer session in scraper profile
    },
    "IT": {
        "sc_base":         "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "sc_display_name": "Italy",
        "sc_mkid":         "APJ6JRA9NG5V4",
        "amazon_home":     "https://www.amazon.it/",
        "review_url":      "https://www.amazon.it/gp/customer-reviews/",
        "country":         "IT",
        "image_fetch":     False,  # no amazon.it customer session in scraper profile
    },
    "ES": {
        "sc_base":         "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "sc_display_name": "Spain",
        "sc_mkid":         "A1RKKUPIHCS9HS",
        "amazon_home":     "https://www.amazon.es/",
        "review_url":      "https://www.amazon.es/gp/customer-reviews/",
        "country":         "ES",
        "image_fetch":     False,  # no amazon.es customer session in scraper profile
    },
    "JP": {
        "sc_base":     "https://sellercentral.amazon.co.jp/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.co.jp/",
        "review_url":  "https://www.amazon.co.jp/gp/customer-reviews/",
        "country":     "JP",
    },
    "IN": {
        "sc_base":     "https://sellercentral.amazon.in/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.in/",
        "review_url":  "https://www.amazon.in/gp/customer-reviews/",
        "country":     "IN",
    },
}

# ═══════════════════════════════════════════════════════════════════════════════
# DETECTION AVOIDANCE PROFILES
# ═══════════════════════════════════════════════════════════════════════════════

_PROFILES = {
    "LOW": {
        "nav_delay":    (0.5,  1.5),
        "read_delay":   (0.3,  0.8),
        "batch_delay":  (0.5,  1.5),
        "fetch_jitter": (0,    150),
        "batch_min":    20,
        "batch_max":    30,
        "scroll":       False,
    },
    "MEDIUM": {
        "nav_delay":    (2.0,  5.0),
        "read_delay":   (1.0,  2.5),
        "batch_delay":  (2.0,  4.5),
        "fetch_jitter": (0,    600),
        "batch_min":    15,
        "batch_max":    22,
        "scroll":       True,
    },
    "HIGH": {
        "nav_delay":    (4.0, 10.0),
        "read_delay":   (2.0,  5.0),
        "batch_delay":  (5.0, 12.0),
        "fetch_jitter": (0,   1200),
        "batch_min":    8,
        "batch_max":    15,
        "scroll":       True,
    },
}

# ═══════════════════════════════════════════════════════════════════════════════
# INTERNALS — no changes needed below this line
# ═══════════════════════════════════════════════════════════════════════════════

ALL_HEADERS = [
    'ASIN', 'Created 날짜', '사진 유무', 'Reviewer', 'Review Ratings',
    'Review Title', '본문', 'Product Rating', 'Ratings Count',
    'Domain Code', '국가', 'Review Link', 'Image URL', 'Review ID',
    'Order ID',   # populated only when FETCH_ORDER_ID = True
]
IDX = {h: i for i, h in enumerate(ALL_HEADERS)}


def _normalize_date(s):
    """Normalize SC date strings to yyyy-mm-dd. Falls back to original on parse failure."""
    for fmt in ("%B %d, %Y", "%d %B %Y"):
        try:
            return datetime.strptime(s.strip(), fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return s


def _find_stale_row_ids(rows):
    """Return set of Review IDs whose Reviewer+Title are identical to the immediately preceding row.

    Different Review IDs sharing the same Reviewer AND Title in consecutive positions are
    a signature of DOM recycling (a card's lazy-loaded content hasn't updated yet).
    """
    stale = set()
    for i in range(1, len(rows)):
        prev, cur = rows[i - 1], rows[i]
        if (cur[IDX['Review ID']] != prev[IDX['Review ID']]
                and cur[IDX['Reviewer']]
                and cur[IDX['Reviewer']] == prev[IDX['Reviewer']]
                and cur[IDX['Review Title']]
                and cur[IDX['Review Title']] == prev[IDX['Review Title']]):
            stale.add(cur[IDX['Review ID']])
    return stale


def _make_extract_js(domain_code, country):
    return f"""
() => {{
  const cards = document.querySelectorAll('.reviewContainer[data-testid]');
  const rows = [];
  cards.forEach(card => {{
    const reviewId    = card.getAttribute('data-testid').replace('review-', '');
    const rating      = card.querySelector('kat-star-rating.reviewRating')?.getAttribute('value') || '';
    const rdText      = card.querySelector('.css-g7g1lz')?.textContent?.trim() || '';
    const rdMatch     = rdText.match(/^Review by (.+?) on (.+)$/);
    const reviewer    = rdMatch?.[1] || '';
    const createdDate = rdMatch?.[2] || '';
    const titleEl     = card.querySelector('#' + reviewId + '-title');
    const title       = titleEl?.querySelector('b')?.textContent?.trim() || titleEl?.textContent?.trim() || '';
    const body        = (document.getElementById('review-content-' + reviewId)?.innerText || '').trim().replace(/\\n/g,' ');
    const reviewLink  = card.querySelector('kat-link[href*="customer-reviews/' + reviewId + '"]')?.getAttribute('href') || '';
    // Collect all meta label→value pairs from the review card detail rows.
    // Prefer 'Child ASIN'; fall back to 'ASIN' only when no child entry exists.
    // Also try extracting ASIN from the product link href (most reliable).
    let childAsin = '';
    const _meta = {{}};
    card.querySelectorAll('.css-yyccc7').forEach(r => {{
      const label = r.querySelector('.css-1ggdaz4')?.textContent?.trim();
      const divs  = r.querySelectorAll('div');
      const val   = [...divs].slice(1).map(d => d.textContent.trim()).filter(Boolean)[0] || '';
      if (label) _meta[label] = val;
    }});
    if (_meta['Child ASIN']) {{
      childAsin = _meta['Child ASIN'];
    }} else {{
      // Try extracting from the product link kat-link href (contains /dp/ASIN or ?ASIN=)
      const prodLink = card.querySelector('kat-link[href*="/dp/"], kat-link[href*="ASIN="]')?.getAttribute('href') || '';
      const asinFromLink = prodLink.match(/\\/dp\\/([A-Z0-9]{{10}})/)?.[1]
                        || new URLSearchParams(prodLink.split('?')[1] || '').get('ASIN') || '';
      childAsin = asinFromLink || _meta['ASIN'] || '';
    }}
    const pStarEl       = card.querySelector('.asinDetail kat-star-rating');
    const productRating = pStarEl?.getAttribute('value') || '';
    const ratingsCount  = pStarEl?.getAttribute('review') || '';
    // Derive actual domain/country from review link — more reliable than the
    // scraper-assigned value when a marketplace switch silently falls back.
    let domainCode = '{domain_code}';
    let countryCode = '{country}';
    if (reviewLink) {{
      const m = reviewLink.match(/amazon\\.(com(?!\\.)|co\\.uk|co\\.jp|de|fr|it|es|in)/);
      if (m) {{
        const _map = {{'com':'US','co.uk':'UK','co.jp':'JP','de':'DE','fr':'FR','it':'IT','es':'ES','in':'IN'}};
        if (_map[m[1]]) {{ domainCode = _map[m[1]]; countryCode = _map[m[1]]; }}
      }}
    }}
    rows.push([childAsin, createdDate, 'N', reviewer, rating, title, body,
               productRating, ratingsCount, domainCode, countryCode, reviewLink, '', reviewId, '']);
  }});
  return rows;
}}
"""


SCROLL_JS = """
() => new Promise(resolve => {
  const maxScroll = document.body.scrollHeight;
  let pos = 0;
  const cap = setTimeout(resolve, 8000);  // hard cap: always resolves within 8 s
  const tick = () => {
    pos += Math.random() * 180 + 60;
    window.scrollTo(0, Math.min(pos, maxScroll));
    if (pos < maxScroll) setTimeout(tick, Math.random() * 120 + 60);
    else { clearTimeout(cap); resolve(); }
  };
  tick();
})
"""

# Waits until every visible review card has loaded its reviewer text.
# Guards against lazy-rendered cards that still show stale content from a recycled DOM slot.
_CONTENT_READY_JS = """
() => {
    const cards = document.querySelectorAll('.reviewContainer[data-testid]');
    if (!cards.length) return false;
    return [...cards].every(c =>
        (c.querySelector('.css-g7g1lz')?.textContent?.trim() || '').length > 0
    );
}
"""


def _make_batch_fetch_js(review_url_base):
    return f"""
async (args) => {{
  const reviewIds = args[0];
  const jitterMin = args[1];
  const jitterMax = args[2];
  const results = {{}};
  await Promise.all(reviewIds.map(async (id, idx) => {{
    await new Promise(r => setTimeout(r, idx * (Math.random() * (jitterMax - jitterMin) + jitterMin)));
    try {{
      const resp = await fetch('{review_url_base}' + id, {{
        credentials: 'include',
        headers: {{
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
          'Accept-Language': 'en-US,en;q=0.9',
          'Accept-Encoding': 'gzip, deflate, br',
          'Referer': '{review_url_base}',
          'Sec-Fetch-Dest': 'document',
          'Sec-Fetch-Mode': 'navigate',
          'Sec-Fetch-Site': 'same-origin',
          'Upgrade-Insecure-Requests': '1'
        }}
      }});
      const html = await resp.text();
      const doc  = new DOMParser().parseFromString(html, 'text/html');
      const tiles = doc.querySelectorAll('[data-hook="review-image-tile"]');
      results[id] = [...tiles].map(el => {{
        const src = el.getAttribute('src') || el.querySelector('img')?.getAttribute('src') || '';
        return src.replace(/\\._[A-Z0-9_,]+_\\./, '.');
      }}).filter(Boolean);
    }} catch(e) {{
      results[id] = [];
    }}
  }}));
  return results;
}}
"""


async def simulate_reading(page, prof):
    if prof["scroll"]:
        try:
            await page.evaluate(SCROLL_JS)
        except Exception:
            pass
    await asyncio.sleep(random.uniform(*prof["read_delay"]))


def _out_file(domain):
    return os.path.join(OUT_DIR, f"{domain}_seller_central_reviews.csv")


def _csv_write_header(path, headers):
    """Create/overwrite CSV with header row only."""
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        csv.writer(f).writerow(headers)


def _csv_append_rows(path, rows):
    """Append rows to an existing CSV (no header)."""
    with open(path, 'a', newline='', encoding='utf-8-sig') as f:
        csv.writer(f).writerows(rows)


def _csv_rewrite(path, headers, rows):
    """Full rewrite — used at end to apply image data to all rows."""
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)


def _apply_column_filter(path):
    """Trim a finished CSV down to HEADERS_TO_INCLUDE, if set.

    Call this only once a domain's file is fully done (all EU sub-countries +
    image enrichment complete) — scrape_domain always writes the full
    ALL_HEADERS layout internally so the shared EU CSV stays in a format the
    next sub-country's read-and-append step can index into.
    """
    if not HEADERS_TO_INCLUDE or not os.path.exists(path):
        return
    with open(path, encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        header = next(reader, None)
        rows = list(reader)
    if not header:
        return
    idx = {h: i for i, h in enumerate(header)}
    keep_idx = [idx[h] for h in HEADERS_TO_INCLUDE if h in idx]
    _csv_rewrite(path, [header[k] for k in keep_idx], [[row[k] for k in keep_idx] for row in rows])


async def _switch_sc_marketplace(page, display_name, prof):
    """Switch SC Europe to a specific marketplace via the two-level account switcher.

    Flow:
      1. Navigate to /home (dropdown only renders account items here).
      2. Click header to open the account list.
      3. Click "Spigen EU" — this EXPANDS its country sub-list (no navigation).
      4. Click the target country item (indented) — this DOES navigate.
      5. Capture mons_sel_* URL params, then navigate to reviews with target mkid.
    """
    print(f"  Switching SC marketplace → {display_name} ...", end=" ", flush=True)
    from urllib.parse import urlparse, parse_qs
    target_dc = next((v for v in _DOMAINS.values()
                      if v.get("sc_display_name") == display_name), None)
    target_mkid = target_dc.get("sc_mkid", "") if target_dc else ""
    try:
        await page.goto("https://sellercentral-europe.amazon.com/home",
                        wait_until="domcontentloaded", timeout=30000)

        # Guard: if session expired, wait for re-login then retry (up to 3 attempts).
        for _attempt in range(3):
            if not any(x in page.url for x in ["/ap/", "signin", "mfa"]):
                break
            print(f"\n  ⚠  SC Europe session expired before switching to {display_name}.")
            if _creds and await ensure_logged_in(page, "EU", _creds, screenshot_dir=SCREENSHOT_DIR):
                pass  # automated login succeeded — fall through to retry the goto below
            else:
                print(f"  Complete login + OTP in Chrome — scraper will retry automatically.")
                if sys.stdin.isatty():
                    input(f"  Press Enter after logging in... ")
                else:
                    for _s in range(MID_RUN_LOGIN_WAIT_SECONDS, 0, -1):
                        print(f"  {_s}s remaining …  ", end="\r", flush=True)
                        await asyncio.sleep(1)
                    print()
            await page.goto("https://sellercentral-europe.amazon.com/home",
                            wait_until="domcontentloaded", timeout=30000)

        await page.wait_for_selector('.dropdown-account-switcher-header', timeout=15000)
        await asyncio.sleep(0.5)

        # Run the full dropdown interaction in a single JS evaluate to avoid CDP
        # round-trip delays that let the dropdown auto-close between steps.
        # Sequence: click header → wait for Spigen EU → click Spigen EU →
        # wait for country sub-list → click target country (triggers navigation).
        # The evaluate will throw "context destroyed" when navigation fires — that's expected.
        try:
            async with page.expect_navigation(timeout=15000):
                try:
                    await page.evaluate(f"""
                        async () => {{
                            const header = document.querySelector('.dropdown-account-switcher-header');
                            if (!header) throw new Error('no header');
                            header.click();

                            const waitFor = (selector, text, maxMs=5000) => new Promise((resolve, reject) => {{
                                const start = Date.now();
                                const iv = setInterval(() => {{
                                    const el = [...document.querySelectorAll(selector)]
                                        .find(e => e.textContent.trim() === text);
                                    if (el) {{ clearInterval(iv); resolve(el); return; }}
                                    if (Date.now() - start > maxMs) {{ clearInterval(iv); reject(new Error('timeout: ' + text)); }}
                                }}, 80);
                            }});

                            const spigen = await waitFor('.dropdown-account-switcher-list-item-label', 'Spigen EU');
                            spigen.click();

                            const country = await waitFor('.dropdown-account-switcher-list-item-indented', '{display_name}');
                            country.click();
                        }}
                    """)
                except Exception:
                    pass  # "Execution context destroyed" when navigation fires — expected
        except Exception as e:
            print(f"WARN: country navigation failed ({e}) — falling back to mkid-only")
            return f"mons_sel_mkid={target_mkid}" if target_mkid else ""
        await asyncio.sleep(0.3)

        # Capture mons_sel_* params from the landed URL.
        # Use the exact mkid from the URL (includes amzn1.mp.o. prefix SC requires).
        qs = parse_qs(urlparse(page.url).query, keep_blank_values=True)
        dir_mcid  = qs.get("mons_sel_dir_mcid", [""])[0]
        mkid_from_url = qs.get("mons_sel_mkid",     [""])[0]

        if dir_mcid and mkid_from_url:
            mons_params = f"mons_sel_dir_mcid={dir_mcid}&mons_sel_mkid={mkid_from_url}"
            await page.goto(
                f"https://sellercentral-europe.amazon.com/brand-customer-reviews/?{mons_params}",
                wait_until="domcontentloaded", timeout=30000,
            )
            await asyncio.sleep(random.uniform(*prof["read_delay"]))
            print("done")
            return mons_params

        result_params = "&".join(f"{k}={v[0]}" for k, v in qs.items() if k.startswith("mons_sel"))
        if result_params:
            print(f"done (params: {result_params})")
            return result_params

        if target_mkid:
            print("done (mkid-only fallback)")
            return f"mons_sel_mkid={target_mkid}"

        print("WARN: could not determine marketplace params")
        return ""

    except Exception as e:
        print(f"WARN: marketplace switch failed ({e})")
        return f"mons_sel_mkid={target_mkid}" if target_mkid else ""


# Each of the 4 login credentials is a genuinely separate Amazon account
# identity (confirmed via distinct CIDs in diagnostic testing), each with its
# own delegated "full-page-account-switcher" tree of business entities and,
# beneath each entity, specific country/marketplace registrations. On a
# fresh/shared Chrome profile, navigating to one domain's URL while a
# DIFFERENT domain's identity is the currently-active session still passes
# the crude "not on a signin page" check — Amazon serves that OTHER identity's
# account picker instead of prompting a real login. _resolve_account() closes
# that gap: it verifies the correct identity is active (not just *an*
# identity) and, if not, forces a logout + genuine re-login with this
# domain's own credentials before selecting the actual business + country.
_TARGET_TOP_LEVEL = {"US": "Spigen Inc", "EU": "Spigen EU", "JP": "Spigen 公式直営店", "IN": "Spigen India"}
_TARGET_SUB_LEVEL = {"US": "United States", "EU": "Germany", "JP": "Japan", "IN": "India"}
# EU only needs an anchor country here (Germany, matching the script's
# existing "DE scraped first" assumption) — the pre-existing
# _switch_sc_marketplace() dropdown mechanism, which only works once already
# inside a resolved account, handles switching to the other EU sub-countries
# during the actual scrape loop; unchanged, no other code path is affected.


async def _resolve_account(page, domain, creds, url, *, max_attempts=2):
    """Ensure the correct Amazon business/marketplace identity is active for
    `domain` on the full-page-account-switcher picker (confirmed via live
    diagnostic testing — see git history for the captured HTML evidence).
    Returns True once resolved (or if the picker isn't showing at all —
    already past it), False if unresolved after retries."""
    top_name = _TARGET_TOP_LEVEL.get(domain)
    if not top_name:
        return True
    sub_name = _TARGET_SUB_LEVEL.get(domain)
    label_sel = ".full-page-account-switcher-account-label"

    for attempt in range(1, max_attempts + 1):
        try:
            picker_showing = await page.get_by_text("Select an account", exact=False).count() > 0
        except Exception:
            picker_showing = False
        if not picker_showing:
            return True

        target_present = await page.locator(f"{label_sel}:text-is('{top_name}')").count() > 0
        if not target_present:
            print(f"  [{domain}] wrong identity active (expected '{top_name}' not in picker, attempt {attempt}) — logging out and re-authenticating")
            try:
                from urllib.parse import urlparse
                origin = urlparse(page.url)
                await page.goto(f"{origin.scheme}://{origin.netloc}/sign-out", wait_until="domcontentloaded", timeout=30000)
            except Exception as e:
                print(f"  [{domain}] sign-out navigation failed: {e}")
            if creds.get(domain):
                await ensure_logged_in(page, domain, creds, screenshot_dir=SCREENSHOT_DIR)
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            except Exception:
                pass
            continue

        try:
            await page.locator(f"{label_sel}:text-is('{top_name}')").first.click()
            await asyncio.sleep(1.2)
            if sub_name:
                sub_loc = page.locator(f"{label_sel}:text-is('{sub_name}')").first
                if await sub_loc.count():
                    await sub_loc.click()
                    await asyncio.sleep(0.8)
            confirm = page.locator("[data-test='confirm-selection']").first
            if await confirm.count():
                await confirm.click()
                await page.wait_for_load_state("domcontentloaded", timeout=15000)
                await asyncio.sleep(1.5)
            print(f"  [{domain}] account resolved: {top_name}" + (f" / {sub_name}" if sub_name else ""))
            return True
        except Exception as e:
            print(f"  [{domain}] account selection failed (attempt {attempt}): {e}")

    return False


async def _enrich_rows_with_images(all_rows, dc, page, prof, domain=None):
    """Fetch reviewer images for rows from a single domain. Enriches rows in-place."""
    if not all_rows:
        return 0
    fetch_js  = _make_batch_fetch_js(dc["review_url"])
    jitter    = prof["fetch_jitter"]
    batch_min = prof["batch_min"]
    batch_max = prof["batch_max"]

    print(f"  Switching to {dc['amazon_home']} for image fetching …")
    try:
        await page.goto(dc["amazon_home"], wait_until="domcontentloaded", timeout=30000)
    except Exception as e:
        # Image fetch is enrichment, not core data — a flaky connection here
        # shouldn't lose the review rows already scraped for this domain.
        print(f"  WARN could not reach {dc['amazon_home']} for image fetch ({e}) — skipping images for this domain")
        return 0
    await asyncio.sleep(random.uniform(1.5, 3.0))

    if _creds and domain:
        from sc_auth import ensure_customer_logged_in
        await ensure_customer_logged_in(page, domain, _creds, screenshot_dir=SCREENSHOT_DIR)

    review_ids = [row[IDX['Review ID']] for row in all_rows]
    id_to_row  = {row[IDX['Review ID']]: row for row in all_rows}
    total      = 0
    i          = 0

    print(f"  Image fetch     : batches {batch_min}–{batch_max}, jitter {jitter[0]}–{jitter[1]}ms")
    while i < len(review_ids):
        batch = review_ids[i:i + random.randint(batch_min, batch_max)]
        try:
            results = await page.evaluate(fetch_js, [batch, jitter[0], jitter[1]])
        except Exception as _err:
            if "closed" in str(_err).lower():
                raise
            print(f"  WARN image batch failed ({_err}) — re-navigating and retrying …")
            try:
                await page.goto(dc["amazon_home"], wait_until="domcontentloaded", timeout=30000)
                await asyncio.sleep(random.uniform(1.5, 3.0))
                results = await page.evaluate(fetch_js, [batch, jitter[0], jitter[1]])
            except Exception:
                results = {}
        for rid, imgs in results.items():
            if imgs and rid in id_to_row:
                id_to_row[rid][IDX['사진 유무']] = 'Y'
                id_to_row[rid][IDX['Image URL']] = '|'.join(imgs)
                total += 1
        i   += len(batch)
        done = min(i, len(review_ids))
        print(f"    {done}/{len(review_ids)}  ({total} with images)")
        if done < len(review_ids):
            await asyncio.sleep(random.uniform(*prof["batch_delay"]))
    return total


async def _enrich_csv_with_images(csv_path, page, prof):
    """Read csv_path, fetch reviewer images grouped by Domain Code, rewrite the file.

    Used by the EU flow to run image fetching for all countries in one pass
    after all page scraping is complete.
    """
    if not os.path.exists(csv_path):
        return 0

    with open(csv_path, encoding='utf-8-sig') as f:
        reader  = csv.reader(f)
        headers = next(reader, None)
        rows    = list(reader)

    if not headers or not rows:
        return 0

    file_idx = {h: idx for idx, h in enumerate(headers)}
    if 'Domain Code' not in file_idx or 'Review ID' not in file_idx:
        print(f"  SKIP image fetch — Domain Code / Review ID column missing in {csv_path}")
        return 0

    dc_col    = file_idx['Domain Code']
    rid_col   = file_idx['Review ID']
    img_col   = file_idx.get('Image URL')
    photo_col = file_idx.get('사진 유무')

    by_domain = defaultdict(list)
    for i, row in enumerate(rows):
        by_domain[row[dc_col]].append(i)

    jitter    = prof["fetch_jitter"]
    batch_min = prof["batch_min"]
    batch_max = prof["batch_max"]
    total     = 0

    for dc_code, row_indices in by_domain.items():
        if dc_code not in _DOMAINS:
            print(f"  SKIP [{dc_code}] — not in domain registry")
            continue
        if not _creds and not _DOMAINS[dc_code].get("image_fetch", True):
            # On the Mac, image_fetch:False is a hard skip — no automated way
            # to establish a customer session there, so don't waste time.
            # On Apify (_creds set), always attempt it — ensure_customer_logged_in
            # below will actually establish the session this flag assumes exists.
            print(f"  SKIP [{dc_code}] — no customer session for image fetch (amazon.{dc_code.lower()} not logged in)")
            continue
        dc       = _DOMAINS[dc_code]
        fetch_js = _make_batch_fetch_js(dc["review_url"])

        print(f"\n  Image fetch [{dc_code}] : navigating to {dc['amazon_home']} …")
        try:
            await page.goto(dc["amazon_home"], wait_until="domcontentloaded", timeout=30000)
        except Exception as e:
            # Image fetch is enrichment, not core data — a flaky connection
            # here (seen: net::ERR_CONNECTION_RESET) shouldn't blow away the
            # review rows already scraped for every other EU country.
            print(f"  WARN [{dc_code}] could not reach {dc['amazon_home']} for image fetch ({e}) — skipping this country's images")
            continue
        await asyncio.sleep(random.uniform(1.5, 3.0))

        if _creds:
            from sc_auth import ensure_customer_logged_in
            await ensure_customer_logged_in(page, dc_code, _creds, screenshot_dir=SCREENSHOT_DIR)

        review_ids = [rows[i][rid_col] for i in row_indices]
        id_to_idx  = {rows[i][rid_col]: i for i in row_indices}

        print(f"  Image fetch     : {len(review_ids)} reviews, batches {batch_min}–{batch_max}, jitter {jitter[0]}–{jitter[1]}ms")
        i = 0
        while i < len(review_ids):
            batch = review_ids[i:i + random.randint(batch_min, batch_max)]
            try:
                results = await page.evaluate(fetch_js, [batch, jitter[0], jitter[1]])
            except Exception as _err:
                if "closed" in str(_err).lower():
                    raise
                print(f"  WARN image batch failed ({_err}) — re-navigating and retrying …")
                try:
                    await page.goto(dc["amazon_home"], wait_until="domcontentloaded", timeout=30000)
                    await asyncio.sleep(random.uniform(1.5, 3.0))
                    results = await page.evaluate(fetch_js, [batch, jitter[0], jitter[1]])
                except Exception:
                    results = {}
            for rid, imgs in results.items():
                if imgs and rid in id_to_idx:
                    row_i = id_to_idx[rid]
                    if photo_col is not None:
                        rows[row_i][photo_col] = 'Y'
                    if img_col is not None:
                        rows[row_i][img_col] = '|'.join(imgs)
                    total += 1
            i   += len(batch)
            done = min(i, len(review_ids))
            print(f"    [{dc_code}] {done}/{len(review_ids)}  ({total} total with images)")
            if done < len(review_ids):
                await asyncio.sleep(random.uniform(*prof["batch_delay"]))

        # Flush after each country so a crash mid-run doesn't lose already-fetched images.
        _csv_rewrite(csv_path, headers, rows)
        print(f"    [{dc_code}] images saved to CSV")
    return total


_ORDER_ID_FETCH_JS = """async (url) => {
    const res = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/json' } });
    const text = await res.text();
    if (!res.ok) return { __error: `HTTP ${res.status}`, __body: text.slice(0, 200) };
    try { return JSON.parse(text); }
    catch (e) { return { __error: `parse: ${e.message}`, __body: text.slice(0, 200) }; }
}"""


async def _fetch_order_id_map(page, domain, mons_params=""):
    """Call the SC internal reviews API and return {reviewId: orderId} for all pages.

    The API is the same endpoint the brand-customer-reviews UI uses internally.
    It returns orderId directly per review (no SP-API needed).
    ~84 % of EU/US reviews have an orderId; non-verified purchases return "".

    Runs the fetch() as in-page JS on the already-authenticated Playwright tab
    (same-origin, real browser cookies/headers) rather than a standalone Python
    requests.Session — a prior version re-extracted cookies via ctx.cookies()
    filtered by hostname substring, which silently dropped Amazon's root-domain
    auth cookies (session-id, at-main, etc. are set on .amazon.com, which doesn't
    contain the sellercentral subdomain as a substring) and caused every call to
    land on an HTML login redirect instead of JSON.
    """
    dc = _DOMAINS[domain]
    # Derive API base: replace the Playwright brand-reviews UI path with the API path
    api_base = dc["sc_base"].replace("/brand-customer-reviews/", "/brandcustomerreviews/api/reviews")

    order_map  = {}
    page_id    = 0
    page_size  = 50
    total_pages = 1

    while page_id < total_pages:
        qs = (f"?pageId={page_id}&pageSize={page_size}"
              f"&sortByType=REVIEW_CREATED_DATE&isAscending=false&includeDone=false")
        if mons_params:
            qs += f"&{mons_params}"
        url = api_base + qs
        try:
            data = await page.evaluate(_ORDER_ID_FETCH_JS, url)
        except Exception as e:
            print(f"  WARN order ID API error page {page_id}: {e}")
            break

        if isinstance(data, dict) and data.get("__error"):
            print(f"  WARN order ID API error page {page_id}: {data['__error']}  body: {data.get('__body','')!r}")
            break

        total_pages = data.get("totalPageCount", 1)
        for rev in data.get("reviews", []):
            rid = rev.get("reviewId", "")
            oid = rev.get("orderId") or ""
            if rid:
                order_map[rid] = oid

        page_id += 1
        if page_id < total_pages:
            await asyncio.sleep(0.3)

    return order_map


async def scrape_domain(domain, page, ctx, prof, asin_filter, out_file=None, append=False, pages=None, skip_images=False, mkp_params=""):
    """Scrape one domain end-to-end. Returns (total_rows, total_with_imgs).

    out_file    : override output path (used by EU group to share one CSV).
    append      : skip header write and load existing rows from out_file first
                  (used for EU sub-countries 2-5 so they append to the shared file).
    pages       : page limit for this domain (overrides PAGES global).
    skip_images : defer image fetching to the caller (used by EU so images are
                  fetched for all countries in one pass after all scraping is done).
                  Ignored when FETCH_IMAGES = False.
    mkp_params  : pre-captured mons_sel_* URL params from a prior marketplace switch.
                  When provided, skips the internal dropdown switch and embeds these
                  params in every page URL so the tab stays locked to the right
                  marketplace even if a parallel tab changes the shared session cookie.
    """
    pages      = pages if pages is not None else PAGES
    dc         = _DOMAINS[domain]
    out_file   = out_file or _out_file(domain)
    extract_js = _make_extract_js(domain, dc["country"])
    params     = f"?pageSize={PAGE_SIZE}&stars={STAR_FILTER}"
    if mkp_params:
        params += f"&{mkp_params}"

    print(f"\n{'═'*60}")
    print(f"  Domain : {domain}  ({dc['sc_base']})")
    print(f"  Pages  : {pages}  |  Page size: {PAGE_SIZE}  |  Stars: {STAR_FILTER}  |  Output: {out_file}")
    print(f"{'═'*60}")

    # Switch SC marketplace via UI dropdown if needed.
    # Skipped when mkp_params is already provided by the caller (EU parallel flow).
    if "sc_display_name" in dc and not mkp_params:
        captured = await _switch_sc_marketplace(page, dc["sc_display_name"], prof)
        if captured:
            params += f"&{captured}"

    # ── Step 1: scrape pages — write header once, append after each page ──
    if APPEND_CSV or append:
        # Resume / EU-group append: read existing rows for dedup + image enrichment
        all_rows = []
        if os.path.exists(out_file):
            with open(out_file, encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                next(reader, None)  # skip header
                all_rows = list(reader)
            label = "Resume mode" if APPEND_CSV else "Appending to shared file"
            print(f"  {label}   : loaded {len(all_rows)} existing rows from CSV")
    else:
        _csv_write_header(out_file, ALL_HEADERS)
        all_rows = []
    # Rows already in the CSV before this call — used below so this domain's order-ID
    # pass never clobbers Order IDs a *different* country's pass already wrote into the
    # same shared EU CSV (its order_map only knows its own country's review IDs, so
    # blindly overwriting every row with "" on no-match wiped out earlier countries).
    _preexisting_rids = {row[IDX['Review ID']] for row in all_rows if row[IDX['Review ID']]}

    p = START_PAGE
    _t_retries = 0  # retry counter for transient Playwright race-condition errors
    _consecutive_empty = 0  # counts genuinely-zero-review pages in a row
    while p <= pages:
        url = dc["sc_base"] + params + (f"&pageNumber={p}" if p > 1 else "")
        print(f"  Page {p}/{pages} …", end=" ", flush=True)
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            await page.wait_for_selector('.reviewContainer[data-testid]', timeout=15000)
        except Exception as e:
            if "list.remove" in str(e) and _t_retries < 2:
                _t_retries += 1
                await asyncio.sleep(1)
                continue  # retry same page
            _t_retries = 0
            if "closed" in str(e).lower():
                raise  # browser disconnected — abort this domain immediately
            cur_url = page.url
            if any(x in cur_url for x in ["/ap/", "signin", "mfa"]):
                print(f"LOGIN REDIRECT detected — session expired on page {p}")
                if _creds and await ensure_logged_in(page, domain, _creds, screenshot_dir=SCREENSHOT_DIR):
                    continue  # automated login succeeded — retry same page
                print(f"  ⚠  Complete login + OTP in Chrome now.")
                if sys.stdin.isatty():
                    input(f"  Press Enter after logging in to retry page {p}... ")
                else:
                    print(f"  Waiting {MID_RUN_LOGIN_WAIT_SECONDS}s for login, then retrying page {p}...")
                    for remaining in range(MID_RUN_LOGIN_WAIT_SECONDS, 0, -1):
                        print(f"  {remaining}s remaining …  ", end="\r", flush=True)
                        await asyncio.sleep(1)
                    print()
                continue  # retry same page — do NOT increment p
            # wait_for_selector timing out (not a login redirect, not a browser
            # crash) almost always means this page has zero review cards, i.e.
            # we've run past the last real page of results. Confirm with a
            # direct DOM check, and only treat it as "end of pagination" after
            # two such pages in a row — guards against one-off network blips
            # burning through the rest of a large PAGES budget page-by-page.
            try:
                card_count = await page.evaluate(
                    "document.querySelectorAll('.reviewContainer[data-testid]').length"
                )
            except Exception:
                card_count = -1  # couldn't check — don't treat as confirmed-empty
            if card_count == 0:
                _consecutive_empty += 1
                print(f"SKIP (0 reviews — page {p} likely past end of results, {_consecutive_empty}/2)")
                if _consecutive_empty >= 2:
                    print(f"  ↳ 2 consecutive empty pages — stopping {domain} at page {p} (reached end of available reviews)")
                    break
                p += 1
                continue
            print(f"SKIP (timeout/error: {e})")
            p += 1
            continue
        _t_retries = 0
        _consecutive_empty = 0
        # Verify the account-switcher header actually shows the target marketplace
        # before trusting this page's data. A switch can silently fail (dropdown
        # UI change, stale fallback params) and leave the tab on the wrong country —
        # catch that here instead of burning the remaining page budget on it.
        if p == START_PAGE and "sc_display_name" in dc:
            try:
                header_text = await page.evaluate(
                    "document.querySelector('.dropdown-account-switcher-header')?.textContent?.trim() || ''"
                )
            except Exception:
                header_text = ""
            if header_text and dc["sc_display_name"].lower() not in header_text.lower():
                print(f"\n  ✗ MARKETPLACE MISMATCH: expected '{dc['sc_display_name']}' but page shows "
                      f"'{header_text}' — aborting {domain} rather than scrape the wrong marketplace")
                break
        await simulate_reading(page, prof)
        # Wait for lazy-rendered card content to finish loading before extracting.
        try:
            await page.wait_for_load_state("networkidle", timeout=5000)
        except Exception:
            pass
        try:
            await page.wait_for_function(_CONTENT_READY_JS, timeout=8000)
        except Exception:
            pass
        try:
            rows = await page.evaluate(extract_js)
        except Exception as e:
            if "list.remove" in str(e) and _t_retries < 2:
                _t_retries += 1
                await asyncio.sleep(1)
                continue  # retry same page from goto
            _t_retries = 0
            raise
        di = IDX['Created 날짜']
        for row in rows:
            row[di] = _normalize_date(row[di])
        # Stale DOM detection: if any card's Reviewer+Title matches the previous card's
        # (different IDs), the lazy sub-component hadn't rendered yet. Reload the page
        # once with a longer wait and re-extract. Drop any rows still stale after retry.
        stale_ids = _find_stale_row_ids(rows)
        if stale_ids:
            print(f"  ↻ {len(stale_ids)} stale rows on page {p} — reloading …", end=" ", flush=True)
            await asyncio.sleep(random.uniform(1.0, 2.5))
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=30000)
                await page.wait_for_selector('.reviewContainer[data-testid]', timeout=15000)
                try:
                    await page.wait_for_load_state("networkidle", timeout=8000)
                except Exception:
                    pass
                try:
                    await page.wait_for_function(_CONTENT_READY_JS, timeout=10000)
                except Exception:
                    pass
                retry_rows = await page.evaluate(extract_js)
                for row in retry_rows:
                    row[di] = _normalize_date(row[di])
                rows = retry_rows
                still_stale = _find_stale_row_ids(rows)
                if still_stale:
                    print(f"resolved {len(stale_ids) - len(still_stale)}, dropping {len(still_stale)} unfixable")
                    rows = [r for r in rows if r[IDX['Review ID']] not in still_stale]
                else:
                    print(f"resolved all {len(stale_ids)}")
            except Exception as _retry_err:
                print(f"retry failed ({_retry_err}) — dropping {len(stale_ids)} stale rows")
                rows = [r for r in rows if r[IDX['Review ID']] not in stale_ids]
        print(f"{len(rows)} reviews  →  flushed to CSV")
        all_rows.extend(rows)
        _csv_append_rows(out_file, rows)   # ← incremental write after every page
        p += 1
        if p <= pages:
            await asyncio.sleep(random.uniform(*prof["nav_delay"]))

    print(f"\n  Total collected : {len(all_rows)}")

    # ── Step 2: deduplicate ───────────────────────────────────────────────
    seen, unique = set(), []
    for row in all_rows:
        rid = row[IDX['Review ID']]
        if rid and rid not in seen:
            seen.add(rid); unique.append(row)
    if len(all_rows) - len(unique):
        print(f"  Dedup           : removed {len(all_rows)-len(unique)} duplicates → {len(unique)} unique")
    all_rows = unique

    # ── Step 2.5: order ID enrichment via SC internal API ────────────────────
    if FETCH_ORDER_ID:
        # Scope stats to rows scraped *this* pass — rows carried over from earlier
        # EU sub-countries in the shared CSV were already resolved (or not) in their
        # own pass and won't appear in this country's order_map.
        new_rids = {r[IDX['Review ID']] for r in all_rows
                    if r[IDX['Review ID']] and r[IDX['Review ID']] not in _preexisting_rids}
        total_reviews = len(new_rids)
        print(f"  Fetching order IDs via SC API ({total_reviews} reviews) …", end=" ", flush=True)
        try:
            order_map = await _fetch_order_id_map(page, domain, mkp_params)
            matched = 0
            for row in all_rows:
                rid = row[IDX['Review ID']]
                oid = order_map.get(rid, "")
                if oid:
                    row[IDX['Order ID']] = oid   # never clobber a good value with "" on no-match
                    if rid in new_rids:
                        matched += 1
            print(f"{matched}/{total_reviews} matched ({100*matched/max(total_reviews,1):.0f}%)")
        except Exception as e:
            print(f"WARN: order ID fetch failed ({e}) — column left empty")

    # ── Step 3: image enrichment (skipped for EU — caller runs _enrich_csv_with_images) ──
    total_with_imgs = 0
    if FETCH_IMAGES and not skip_images:
        total_with_imgs = await _enrich_rows_with_images(all_rows, dc, page, prof, domain=domain)

    # ── Step 4: ASIN filter ───────────────────────────────────────────────
    if asin_filter:
        before   = len(all_rows)
        all_rows = [r for r in all_rows if r[IDX['ASIN']] in asin_filter]
        print(f"  ASIN filter     : kept {len(all_rows)}/{before}")

    # ── Step 4.5: min review date filter ────────────────────────────────
    if MIN_REVIEW_DATE:
        before   = len(all_rows)
        all_rows = [r for r in all_rows if r[IDX['Created 날짜']] >= MIN_REVIEW_DATE]
        print(f"  Date filter     : kept {len(all_rows)}/{before} (>= {MIN_REVIEW_DATE})")

    # ── Step 5: final rewrite with image data ─────────────────────────────
    # Always written in the full ALL_HEADERS layout — the shared EU CSV gets
    # re-read by the *next* sub-country's scrape_domain call (see `append`
    # above), and that read path indexes rows via IDX, which assumes every
    # column is present. Trimming columns here would leave the file one
    # country's pass ahead of what the next country's read expects, causing
    # a "list index out of range" crash. Column trimming (HEADERS_TO_INCLUDE)
    # is applied once, after a domain's file is fully finished — see
    # _apply_column_filter, called from main().
    # all_rows already contains all rows (previous countries loaded at start +
    # newly scraped), so no read-back needed even in append mode.
    _csv_rewrite(out_file, ALL_HEADERS, all_rows)
    print(f"\n  ✓ {domain} done — {total_with_imgs}/{len(all_rows)} with images → {out_file}")

    return len(all_rows), total_with_imgs


def _upload_to_sheets(results, run_date):
    """Merge all domain CSVs into the SC_{yymmdd} worksheet on SHEETS_SPREADSHEET_ID,
    creating it fresh if today's sheet doesn't exist yet. On a same-day re-run
    (e.g. hourly scheduling), only rows whose Review ID isn't already in the sheet
    get appended — existing rows are never rewritten or duplicated. Prints
    per-country (국가) new-vs-already-present counts."""
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        import gspread
    except ImportError as e:
        print(f"\n  SKIP sheets upload — missing library: {e}")
        print("  Install with: pip install gspread google-auth")
        return

    token_path = os.path.expanduser("~/.config/gws_shim/token.json")
    if not os.path.exists(token_path):
        print(f"\n  SKIP sheets upload — credentials not found at {token_path}")
        return

    with open(token_path) as _f:
        tok = json.load(_f)

    creds = Credentials(
        token=tok.get("token"),
        refresh_token=tok.get("refresh_token"),
        token_uri=tok.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=tok.get("client_id"),
        client_secret=tok.get("client_secret"),
        scopes=tok.get("scopes"),
    )
    if creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as e:
            print(f"\n  SKIP sheets upload — token refresh failed: {e}")
            return

    try:
        gc = gspread.authorize(creds)
        spreadsheet = gc.open_by_key(SHEETS_SPREADSHEET_ID)
    except Exception as e:
        print(f"\n  SKIP sheets upload — cannot open spreadsheet: {e}")
        return

    print(f"\n{'═'*60}")
    print(f"  Google Sheets upload  (run date: {run_date})")
    print(f"{'═'*60}")

    header = None
    per_domain_rows = []  # [(domain, [row, ...]), ...] — header trimming/order
                          # already applied upstream by _apply_column_filter.
    for domain, n_rows, n_imgs, status in results:
        if status != "OK" or n_rows == 0:
            print(f"  SKIP [{domain}] — {status}")
            continue
        csv_path = os.path.join(OUT_DIR, f"{domain}_seller_central_reviews.csv")
        if not os.path.exists(csv_path):
            print(f"  SKIP [{domain}] — CSV not found at {csv_path}")
            continue

        with open(csv_path, encoding='utf-8-sig') as _f:
            data = list(csv.reader(_f))

        if not data:
            print(f"  SKIP [{domain}] — CSV is empty")
            continue

        if header is None:
            header = data[0]
        per_domain_rows.append((domain, data[1:]))

    if header is None or not per_domain_rows:
        print("  SKIP sheets upload — no data collected from any domain")
        return []

    rid_idx = header.index("Review ID") if "Review ID" in header else None
    if rid_idx is None:
        print("  WARNING: no 'Review ID' column in HEADERS_TO_INCLUDE — cannot "
              "dedup against existing rows, every scraped row will be treated as new.")
    country_idx = header.index("국가") if "국가" in header else None

    sheet_name = f"SC_{run_date}"
    existing_ws = next((ws_ for ws_ in spreadsheet.worksheets() if ws_.title == sheet_name), None)

    existing_ids = set()
    sheet_header = header
    if existing_ws is not None:
        existing_values = existing_ws.get_all_values()
        if existing_values:
            sheet_header = existing_values[0]
            if "Review ID" in sheet_header:
                existing_rid_idx = sheet_header.index("Review ID")
                existing_ids = {row[existing_rid_idx] for row in existing_values[1:] if len(row) > existing_rid_idx}

    # Reorder each new row into the existing sheet's column order — defensive
    # against HEADERS_TO_INCLUDE having changed since the sheet was first created.
    if sheet_header != header:
        col_map = [header.index(h) if h in header else None for h in sheet_header]
        _reorder = lambda row: [row[i] if i is not None and i < len(row) else "" for i in col_map]
    else:
        _reorder = lambda row: row

    combined_new_rows = []
    seen_this_run = set()
    new_by_country, dup_by_country = defaultdict(int), defaultdict(int)
    for domain, rows in per_domain_rows:
        for row in rows:
            rid = row[rid_idx] if rid_idx is not None and rid_idx < len(row) else None
            country = row[country_idx] if country_idx is not None and country_idx < len(row) else domain
            if rid is not None and (rid in existing_ids or rid in seen_this_run):
                dup_by_country[country] += 1
                continue
            if rid is not None:
                seen_this_run.add(rid)
            combined_new_rows.append(_reorder(row))
            new_by_country[country] += 1

    print()
    for country in sorted(set(new_by_country) | set(dup_by_country)):
        print(f"  [{country}] new {new_by_country.get(country, 0)}  |  already in sheet {dup_by_country.get(country, 0)}")

    if not combined_new_rows:
        print(f"\n  ✓ no new reviews this run — '{sheet_name}' unchanged")
        return []

    try:
        if existing_ws is None:
            ws = spreadsheet.add_worksheet(
                title=sheet_name,
                rows=max(len(combined_new_rows) + 1, 2),
                cols=max(len(sheet_header), 1),
            )
            ws.update(range_name='A1', values=[sheet_header] + combined_new_rows, value_input_option='RAW')
            print(f"\n  ✓ created '{sheet_name}' with {len(combined_new_rows)} rows")
        else:
            existing_ws.append_rows(combined_new_rows, value_input_option='RAW')
            print(f"\n  ✓ appended {len(combined_new_rows)} new rows → '{sheet_name}'")
    except Exception as e:
        print(f"  ✗ upload failed: {e}")

    combined_rows = [sheet_header] + combined_new_rows
    # Returned regardless of whether the Sheets write above succeeded — the
    # scrape itself succeeded either way, and main.py (Apify deployment only)
    # uses this to push the exact same rows to the Actor's default dataset,
    # so the data is downloadable from the Runs/Output tab too.
    return combined_rows


async def main():
    if DETECTION_AVOIDANCE not in _PROFILES:
        raise ValueError("DETECTION_AVOIDANCE must be LOW, MEDIUM, or HIGH.")
    unknown = [d for d in DOMAINS if d not in _DOMAINS and d != "EU"]
    if unknown:
        raise ValueError(f"Unknown domain(s): {unknown}. Choose from: EU | {list(_DOMAINS)}")

    KST = timezone(timedelta(hours=9))
    run_date = datetime.now(KST).strftime("%y%m%d")

    prof = _PROFILES[DETECTION_AVOIDANCE]

    asin_filter = None
    if ASIN_FILTER_FILE:
        with open(os.path.expanduser(ASIN_FILTER_FILE), encoding='utf-8') as af:
            asin_filter = {line.strip() for line in af if line.strip()}
        print(f"ASIN filter loaded: {len(asin_filter)} ASINs from {ASIN_FILTER_FILE}")

    if FETCH_IMAGES_ONLY:
        print(f"Mode        : FETCH_IMAGES_ONLY — skipping scrape, re-fetching images on existing CSVs")
    print(f"Domains     : {DOMAINS}  (parallel)")
    print(f"Pages/domain: {PAGES}  |  Overrides: {PAGES_OVERRIDE or 'none'}  |  Page size: {PAGE_SIZE}  |  Stars: {STAR_FILTER}  |  Avoidance: {DETECTION_AVOIDANCE}")
    print(f"Headless    : {HEADLESS}")

    async with async_playwright() as pw:
        if ISOLATED_TEST_DOMAIN:
            # Diagnostic: test whether ISOLATED_TEST_DOMAIN's credentials
            # reach a genuinely different Amazon account than the shared
            # SCRAPER_PROFILE_DIR profile shows. A prior diagnostic pass
            # found that account-picker only ever listed EU's own delegated
            # accounts (PowerArc/Spigen Direct/Spigen EU/Spigen Inc) even
            # when navigating from JP's or IN's tab — suggesting Amazon's
            # cross-domain session cookies made those tabs LOOK logged in
            # without ever actually authenticating with their own distinct
            # credentials. This uses a brand-new, throwaway profile dir
            # (never the shared one) so there's no possibility of inheriting
            # another domain's session — forces a genuine fresh login.
            import shutil as _shutil
            iso_dir = "/tmp/sc_scraper_isolated_test_profile"
            _shutil.rmtree(iso_dir, ignore_errors=True)
            os.makedirs(iso_dir, exist_ok=True)
            print(f"ISOLATED_TEST_DOMAIN={ISOLATED_TEST_DOMAIN} — using throwaway profile {iso_dir}")
            ctx = await pw.chromium.launch_persistent_context(iso_dir, channel="chrome", headless=HEADLESS)
            page = ctx.pages[0] if ctx.pages else await ctx.new_page()
            url = (_DOMAINS[ISOLATED_TEST_DOMAIN]["sc_base"] if ISOLATED_TEST_DOMAIN != "EU"
                   else "https://sellercentral-europe.amazon.com/brand-customer-reviews/")
            await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            logged_in = not any(x in page.url for x in ["/ap/", "signin", "mfa"])
            print(f"  Pre-login check: {'already logged in (unexpected in a fresh profile!)' if logged_in else 'not logged in, proceeding to real login'}")
            if not logged_in and _creds:
                logged_in = await ensure_logged_in(page, ISOLATED_TEST_DOMAIN, _creds, screenshot_dir=SCREENSHOT_DIR)
            os.makedirs(SCREENSHOT_DIR, exist_ok=True)
            try:
                await page.wait_for_load_state("networkidle", timeout=8000)
            except Exception:
                pass
            await asyncio.sleep(3)
            html = await page.content()
            with open(os.path.join(SCREENSHOT_DIR, f"ISOLATED_{ISOLATED_TEST_DOMAIN}_accountpicker.html"), "w", encoding="utf-8") as f:
                f.write(html)
            await page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"ISOLATED_{ISOLATED_TEST_DOMAIN}_landing.png"))
            print(f"  logged_in={logged_in} — HTML + screenshot captured")
            await ctx.close()
            sys.exit(0)

        if DIAGNOSE_CUSTOMER_LOGIN:
            # Diagnostic: test ONLY the customer-storefront login/interstitial
            # flow for one domain, against the REAL persisted per-domain
            # profile (not a throwaway one) — this is exactly what production
            # image-fetch runs against. Skips Seller Central and the entire
            # scrape/upload pipeline so this stays fast and cheap: one page
            # load, one login attempt, a couple of screenshots, then exit.
            from sc_auth import ensure_customer_logged_in, credential_group
            domain = DIAGNOSE_CUSTOMER_LOGIN
            group = credential_group(domain)
            profile_dir = f"{SCRAPER_PROFILE_DIR}_{group}" if _creds else SCRAPER_PROFILE_DIR
            print(f"DIAGNOSE_CUSTOMER_LOGIN={domain} (profile group {group}) — using persisted profile {profile_dir}")
            os.makedirs(profile_dir, exist_ok=True)
            ctx = await pw.chromium.launch_persistent_context(profile_dir, channel="chrome", headless=HEADLESS)
            page = ctx.pages[0] if ctx.pages else await ctx.new_page()
            dc = _DOMAINS[domain]
            os.makedirs(SCREENSHOT_DIR, exist_ok=True)

            await page.goto(dc["amazon_home"], wait_until="domcontentloaded", timeout=30000)
            await asyncio.sleep(1.5)
            await page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"DIAG_{domain}_0_raw_landing.png"))
            with open(os.path.join(SCREENSHOT_DIR, f"DIAG_{domain}_0_raw_landing.html"), "w", encoding="utf-8") as f:
                f.write(await page.content())

            result = False
            if _creds:
                result = await ensure_customer_logged_in(page, domain, _creds, screenshot_dir=SCREENSHOT_DIR)

            await page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"DIAG_{domain}_9_final.png"))
            with open(os.path.join(SCREENSHOT_DIR, f"DIAG_{domain}_9_final.html"), "w", encoding="utf-8") as f:
                f.write(await page.content())

            # Cheapest possible real signal: try fetching images for just
            # ONE known review ID already in this domain's last CSV, if one
            # exists locally — confirms the actual fetch path, not just the
            # login UI state. Skipped (not an error) if no prior CSV exists.
            image_result = None
            csv_path = _out_file(domain if domain != "EU" else "EU")
            try:
                if os.path.exists(csv_path):
                    with open(csv_path, encoding='utf-8-sig') as f:
                        reader = csv.reader(f)
                        header = next(reader, None)
                        first_row = next(reader, None)
                    if header and first_row and 'Review ID' in header:
                        rid = first_row[header.index('Review ID')]
                        fetch_js = _make_batch_fetch_js(dc["review_url"])
                        results = await page.evaluate(fetch_js, [[rid], 0, 50])
                        image_result = results.get(rid, [])
            except Exception as e:
                image_result = f"ERROR: {e}"

            print(f"  DIAGNOSE_CUSTOMER_LOGIN result: login={result} sample_image_fetch={image_result}")
            await ctx.close()
            sys.exit(0)

        # Each of the 4 top-level domains is a genuinely separate Amazon
        # account identity (confirmed via distinct CIDs in diagnostic
        # testing — see _resolve_account()'s docstring). A single shared
        # Chrome profile can only hold ONE of these identities "actively
        # signed in" at a time; switching between domain tabs in one shared
        # profile forced a fresh logout+re-login on every domain, every
        # single run. Each domain group therefore gets its OWN persistent
        # profile directory — sessions for all 4 then persist independently
        # across days, matching the original "log in once" design intent.
        if HEADLESS and shutil.which("pgrep") and __import__("subprocess").run(
                ["pgrep", "-x", "Google Chrome"], capture_output=True).returncode == 0:
            print("⚠  Chrome is open — close it first (Cmd+Q), then press Enter.")
            if sys.stdin.isatty():
                await asyncio.get_event_loop().run_in_executor(None, input)

        # ── Session check: navigate to each SC URL; only prompt login if needed ──
        _login_endpoints = []
        _seen_eu = False
        for _d in DOMAINS:
            if _d == "EU":
                if not _seen_eu:
                    _seen_eu = True
                    _login_endpoints.append(("EU", "https://sellercentral-europe.amazon.com/brand-customer-reviews/"))
            else:
                _login_endpoints.append((_d, _DOMAINS[_d]["sc_base"]))

        # Per-domain-group profile isolation is only needed (and only ever
        # tested) for the automated-login deployment (_creds set) — the
        # Mac's existing manual workflow keeps its original single shared
        # profile untouched, matching months of prior working behavior
        # there. Splitting into 4 profiles unconditionally would have
        # pointed the Mac at brand-new, empty directories instead of its
        # real login history.
        _use_per_domain_profiles = bool(_creds)
        _shared_ctx = None
        if not _use_per_domain_profiles:
            os.makedirs(SCRAPER_PROFILE_DIR, exist_ok=True)
            print(f"Launching Chrome with scraper profile: {SCRAPER_PROFILE_DIR}")
            _shared_ctx = await pw.chromium.launch_persistent_context(
                SCRAPER_PROFILE_DIR, channel="chrome", headless=HEADLESS)

        print("Opening Seller Central tabs "
              + ("(one profile per account) …" if _use_per_domain_profiles else "…"))
        domain_ctxs = {}
        domain_pages = {}
        needs_login = []
        _shared_existing_pages = list(_shared_ctx.pages) if _shared_ctx else []
        for _label, _url in _login_endpoints:
            if _use_per_domain_profiles:
                _profile_dir = f"{SCRAPER_PROFILE_DIR}_{_label}"
                print(f"  [{_label}] profile: {_profile_dir}")
                os.makedirs(_profile_dir, exist_ok=True)
                _ctx = await pw.chromium.launch_persistent_context(_profile_dir, channel="chrome", headless=HEADLESS)
                _p = _ctx.pages[0] if _ctx.pages else await _ctx.new_page()
            else:
                _ctx = _shared_ctx
                _p = _shared_existing_pages.pop(0) if _shared_existing_pages else await _ctx.new_page()
            domain_ctxs[_label] = _ctx
            domain_pages[_label] = _p
            try:
                await _p.goto(_url, wait_until="domcontentloaded", timeout=30000)
                # Login redirects always land on a URL containing /ap/ or signin
                _logged_in = not any(x in _p.url for x in ["/ap/", "signin", "mfa"])
            except Exception:
                _logged_in = False

            if _logged_in and SCREENSHOT_DIR:
                # Diagnostic checkpoint: capture what the "already logged in"
                # landing page actually looks like, after giving any
                # client-side rendering (account/marketplace selector, etc.)
                # extra time to finish — a blank-but-authenticated shell here
                # was the root cause of an earlier 403 "selection_required"
                # failure on every domain in a from-scratch container.
                try:
                    await _p.wait_for_load_state("networkidle", timeout=8000)
                except Exception:
                    pass
                await asyncio.sleep(3)
                try:
                    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
                    await _p.screenshot(
                        path=os.path.join(SCREENSHOT_DIR, f"{int(time.time())}_{_label}_session_check_landing.png"))
                except Exception:
                    pass

            if not _logged_in and _creds:
                _logged_in = await ensure_logged_in(_p, _label, _creds, screenshot_dir=SCREENSHOT_DIR)

            if _logged_in and _creds and not DIAGNOSE_ACCOUNTS:
                # A crude "not on a signin page" pass does NOT guarantee the
                # right account identity is active — see _resolve_account()'s
                # docstring. Always resolve, even when _logged_in was already
                # True from the shared-profile session check.
                if not await _resolve_account(_p, _label, _creds, _url):
                    print(f"  [{_label}] WARNING: could not resolve the correct account — scrape will likely fail")

            if _logged_in and DIAGNOSE_ACCOUNTS:
                os.makedirs(SCREENSHOT_DIR, exist_ok=True)
                try:
                    html = await _p.content()
                    with open(os.path.join(SCREENSHOT_DIR, f"{_label}_accountpicker.html"), "w", encoding="utf-8") as f:
                        f.write(html)
                except Exception as e:
                    print(f"  [{_label}] diagnose: content() failed: {e}")
                # Confirmed page structure (from a prior diagnostic pass): this
                # is a Vue "full-page-account-switcher" tree, not a flat
                # dropdown. Each top-level account row has its own expander;
                # PowerArc (unrelated business, DOM-first) was accidentally
                # expanded before by a generic "first switcher-like element"
                # selector. This time, expand each named top-level account
                # explicitly, by label text, only once (on the JP tab — the
                # tree is identical across all 4 domains, same underlying
                # Amazon identity/CID, confirmed via the earlier landing
                # screenshots all sharing CID A1VXTSRX565TLZ).
                if _label == "JP":
                    label_sel = ".full-page-account-switcher-account-label"
                    for top_name in ["Spigen Direct", "Spigen EU", "Spigen Inc"]:
                        try:
                            row_label = _p.locator(f"{label_sel}:text-is('{top_name}')").first
                            if not await row_label.count():
                                print(f"  [JP] diagnose: top-level row '{top_name}' not found")
                                continue
                            await row_label.click()
                            await asyncio.sleep(1.5)
                            html2 = await _p.content()
                            safe_name = top_name.replace(" ", "_")
                            with open(os.path.join(SCREENSHOT_DIR, f"JP_expand_{safe_name}.html"), "w", encoding="utf-8") as f:
                                f.write(html2)
                            print(f"  [JP] diagnose: expanded '{top_name}'")
                            # If a Japan-named child row appeared (only expected
                            # under Spigen Direct), expand it too for the
                            # third nesting level the user's screenshot showed.
                            jp_child = _p.locator(f"{label_sel}:text-is('Spigen 公式直営店')").first
                            if await jp_child.count():
                                await jp_child.click()
                                await asyncio.sleep(1.5)
                                html3 = await _p.content()
                                with open(os.path.join(SCREENSHOT_DIR, f"JP_expand_{safe_name}_then_SpigenJP.html"), "w", encoding="utf-8") as f:
                                    f.write(html3)
                                print(f"  [JP] diagnose: expanded 'Spigen 公式直営店' under '{top_name}'")
                            # Reload fresh before trying the next top-level
                            # account, so each expansion starts from the same
                            # clean state (Vue may collapse siblings or not —
                            # don't rely on it).
                            await _p.goto(_url, wait_until="domcontentloaded", timeout=30000)
                            await asyncio.sleep(2)
                        except Exception as e:
                            print(f"  [JP] diagnose: expand '{top_name}' failed: {e}")
                print(f"  [{_label}] diagnose: HTML captured")

            # Always bring every tab to front so the user can verify the correct
            # marketplace is selected — even valid sessions may be on the wrong one.
            try:
                await _p.bring_to_front()
            except Exception:
                pass
            if _logged_in:
                print(f"  [{_label}] Session valid — verify correct marketplace")
            else:
                needs_login.append(_label)
                print(f"  [{_label}] Not logged in — log in now")

        if DIAGNOSE_ACCOUNTS:
            print("\n  DIAGNOSE_ACCOUNTS mode — HTML captured for all domains, exiting without scraping.")
            sys.exit(0)

        login_notice = f"Login required for: {needs_login}\n  " if needs_login else ""
        print(f"\n  {login_notice}→ Log in if needed, then navigate each tab to the correct marketplace.")
        print(f"  Press Enter when ready, or wait {LOGIN_WAIT_SECONDS} s for auto-start.")
        if sys.stdin.isatty():
            await asyncio.get_event_loop().run_in_executor(None, input)
        elif not needs_login:
            print("\r  All sessions valid — starting immediately!    ")
        else:
            wait = LOGIN_WAIT_SECONDS
            print(f"  (non-interactive: starting in {wait} s)")
            for remaining in range(wait, 0, -1):
                print(f"\r  {remaining:3d}s remaining …", end="", flush=True)
                await asyncio.sleep(1)
            print("\r  Starting scrape!                    ")

        # ── Run all domains simultaneously ────────────────────────────────────
        async def _run(domain):
            try:
                group = "EU" if domain == "EU" else domain
                page  = domain_pages[group]
                group_ctx = domain_ctxs[group]
                if domain == "EU":
                    eu_file = os.path.join(OUT_DIR, "EU_seller_central_reviews.csv")
                    eu_rows = 0

                    if not FETCH_IMAGES_ONLY:
                        # Phase 1 — Germany (always first if included, clears or appends per APPEND_CSV)
                        if "DE" in EU_COUNTRIES:
                            n_de, _ = await scrape_domain(
                                "DE", page, group_ctx, prof, asin_filter,
                                out_file=eu_file, append=APPEND_CSV,
                                pages=PAGES_OVERRIDE.get("DE", PAGES), skip_images=True
                            )
                            eu_rows += n_de

                        # Phase 2 — remaining EU countries sequentially on the same tab.
                        # Reusing the active tab keeps the SC Europe session alive between
                        # country switches. Opening a new tab per country was prone to
                        # session expiry mid-run when DE took 30-60 min to complete.
                        eu_rest = [s for s in EU_COUNTRIES if s != "DE"]

                        de_was_written = "DE" in EU_COUNTRIES
                        for i, sub in enumerate(eu_rest):
                            # First country appends only if DE already wrote the CSV;
                            # otherwise respect APPEND_CSV (fresh vs resume).
                            first_append = True if de_was_written else (APPEND_CSV if i == 0 else True)
                            mkp = await _switch_sc_marketplace(
                                page, _DOMAINS[sub]["sc_display_name"], prof
                            )
                            await scrape_domain(
                                sub, page, group_ctx, prof, asin_filter,
                                out_file=eu_file, append=first_append,
                                pages=PAGES_OVERRIDE.get(sub, PAGES),
                                skip_images=True, mkp_params=mkp
                            )

                    # Count actual EU rows from CSV (append mode returns cumulative totals)
                    if os.path.exists(eu_file):
                        with open(eu_file, encoding='utf-8-sig') as _f:
                            eu_rows = sum(1 for _ in _f) - 1  # subtract header

                    # Phase 3 — fetch images for all EU rows in one pass
                    eu_imgs = 0
                    if FETCH_IMAGES:
                        print(f"\n{'═'*60}")
                        print(f"  EU image fetch phase  ({eu_rows} reviews across {len(EU_COUNTRIES)} countries)")
                        print(f"{'═'*60}")
                        eu_imgs = await _enrich_csv_with_images(eu_file, page, prof)
                    _apply_column_filter(eu_file)
                    return ("EU", eu_rows, eu_imgs, "OK")
                else:
                    if FETCH_IMAGES_ONLY:
                        return (domain, 0, 0, "SKIP (FETCH_IMAGES_ONLY — non-EU domain)")
                    eff_pages = PAGES_OVERRIDE.get(domain, PAGES)
                    n_rows, n_imgs = await scrape_domain(
                        domain, page, group_ctx, prof, asin_filter, pages=eff_pages)
                    _apply_column_filter(_out_file(domain))
                    return (domain, n_rows, n_imgs, "OK")
            except Exception as e:
                print(f"\n  ✗ {domain} failed: {e}")
                return (domain, 0, 0, f"FAILED: {e}")

        results = await asyncio.gather(*[_run(d) for d in DOMAINS])

        for _ctx in {id(c): c for c in domain_ctxs.values()}.values():
            await _ctx.close()

    print(f"\n{'═'*60}")
    print("  SUMMARY")
    print(f"{'═'*60}")
    for domain, n_rows, n_imgs, status in results:
        print(f"  {domain:4s}  {n_rows:>5} reviews  {n_imgs:>4} with images  [{status}]")
    print(f"{'═'*60}")

    global LAST_COMBINED_ROWS
    if UPLOAD_TO_SHEETS:
        LAST_COMBINED_ROWS = _upload_to_sheets(results, run_date)

    # Set before the possible sys.exit(1) below so main.py (Apify deployment
    # only) can still read whatever WAS collected even on a partial failure —
    # sys.exit raises immediately, so a return value here wouldn't reach a
    # caller on that path, but a module-level var read after catching
    # SystemExit does.
    if any(status.startswith("FAILED") for _, _, _, status in results):
        sys.exit(1)


if __name__ == '__main__':
    asyncio.run(main())

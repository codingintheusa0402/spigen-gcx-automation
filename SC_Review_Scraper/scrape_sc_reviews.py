#!/usr/bin/env python3
"""Scrape reviews from Amazon Seller Central with configurable options.

Edit the USER CONFIG section below, then run:
    python3 scrape_sc_reviews.py
"""

import asyncio, csv, random, os
from playwright.async_api import async_playwright

# ═══════════════════════════════════════════════════════════════════════════════
# USER CONFIG — edit these before each run
# ═══════════════════════════════════════════════════════════════════════════════

DOMAINS = ["EU", "JP", "IN"]
# List of domains to scrape sequentially. Each gets its own CSV file.
# Single domain example : DOMAINS = ["US"]
# Supported             : "US" | "EU" | "UK" | "DE" | "FR" | "IT" | "ES" | "JP" | "IN"
# "EU" automatically scrapes UK + DE + FR + IT + ES in sequence using each
# country's marketplaceId and writes all reviews into one EU_*.csv file.

PAGES = 10
# Max pages to scrape per domain (50 reviews per page).
# US full scrape ≈ 49 pages (~2,449 reviews).

START_PAGE = 1
# Page to start from. Set > 1 to resume a previously interrupted run.
# When resuming, also set APPEND_CSV = True to avoid overwriting saved rows.

APPEND_CSV = False
# False — overwrites the CSV at the start (default, fresh run).
# True  — appends to an existing CSV without rewriting the header.
#          Use together with START_PAGE to resume an interrupted run.

STAR_FILTER = "1,2,3"
# Comma-separated star ratings to include.
# Critical reviews only: "1,2,3"   All reviews: "1,2,3,4,5"

OUT_DIR = os.path.expanduser("~/Desktop")
# Directory where CSVs are saved.
# Each domain is saved as <OUT_DIR>/<DOMAIN>_seller_central_reviews.csv

HEADERS_TO_INCLUDE = None
# Columns to keep in the output CSV. None → all columns (default).
# Example: ['ASIN', 'Reviewer', 'Review Ratings', 'Review Title', '본문', 'Image URL']
# Full list: ASIN | Created 날짜 | 사진 유무 | Reviewer | Review Ratings | Review Title
#            본문 | Product Rating | Ratings Count | Domain Code | 국가
#            Review Link | Image URL | Review ID

ASIN_FILTER_FILE = None
# Path to a .txt file containing one ASIN per line.
# Only reviews whose ASIN matches an entry in this file will be saved to CSV.
# None → save all reviews regardless of ASIN (default).
# Example: ASIN_FILTER_FILE = "/Users/kevinkim/Desktop/target_asins.txt"

DETECTION_AVOIDANCE = "MEDIUM"
# LOW    — short delays, fastest runs, higher detection risk
# MEDIUM — randomized delays + scroll simulation (recommended for daily use)
# HIGH   — aggressive randomization + long delays (safest for large/frequent scrapes)

HEADLESS = False
# False (default) — connects to your running Chrome via CDP (port 9222).
#                   Browser window stays visible; ideal for watching and debugging.
# True            — launches a headless Chromium using your saved Chrome profile so
#                   existing login sessions/cookies are reused. Chrome must be fully
#                   closed before running in headless mode (profile lock conflict).

CHROME_USER_DATA = os.path.expanduser("~/Library/Application Support/Google/Chrome")
# Path to your Chrome user data directory. Used only when HEADLESS = True.
# Mac default : ~/Library/Application Support/Google/Chrome
# Windows     : %LOCALAPPDATA%/Google/Chrome/User Data

# ═══════════════════════════════════════════════════════════════════════════════
# DOMAIN REGISTRY
# Add new marketplaces here as they are onboarded.
# EU domains (UK/DE/FR/IT/ES) share sellercentral-europe.amazon.com but each
# requires its own marketplaceId parameter. When DOMAINS = ["EU"], the scraper
# automatically loops through all EU_COUNTRIES and writes to one EU_*.csv.
# ═══════════════════════════════════════════════════════════════════════════════

# When "EU" is in DOMAINS, these sub-countries are scraped in order.
EU_COUNTRIES = ["UK", "DE", "FR", "IT", "ES"]

_DOMAINS = {
    "US": {
        "sc_base":     "https://sellercentral.amazon.com/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.com/",
        "review_url":  "https://www.amazon.com/gp/customer-reviews/",
        "country":     "US",
    },
    "UK": {
        "sc_base":        "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "marketplace_id": "A1F83G8C2ARO7P",
        "amazon_home":    "https://www.amazon.co.uk/",
        "review_url":     "https://www.amazon.co.uk/gp/customer-reviews/",
        "country":        "UK",
    },
    "DE": {
        "sc_base":        "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "marketplace_id": "A1PA6795UKMFR9",
        "amazon_home":    "https://www.amazon.de/",
        "review_url":     "https://www.amazon.de/gp/customer-reviews/",
        "country":        "DE",
    },
    "FR": {
        "sc_base":        "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "marketplace_id": "A13V1IB3VIYZZH",
        "amazon_home":    "https://www.amazon.fr/",
        "review_url":     "https://www.amazon.fr/gp/customer-reviews/",
        "country":        "FR",
    },
    "IT": {
        "sc_base":        "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "marketplace_id": "APJ6JRA9NG5V4",
        "amazon_home":    "https://www.amazon.it/",
        "review_url":     "https://www.amazon.it/gp/customer-reviews/",
        "country":        "IT",
    },
    "ES": {
        "sc_base":        "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "marketplace_id": "A1RKKUPIHCS9HS",
        "amazon_home":    "https://www.amazon.es/",
        "review_url":     "https://www.amazon.es/gp/customer-reviews/",
        "country":        "ES",
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
]
IDX = {h: i for i, h in enumerate(ALL_HEADERS)}


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
    let childAsin = '';
    card.querySelectorAll('.css-yyccc7').forEach(r => {{
      const label = r.querySelector('.css-1ggdaz4')?.textContent?.trim();
      const val   = r.querySelectorAll('div')[1]?.textContent?.trim();
      if (label === 'Child ASIN') childAsin = val;
    }});
    const pStarEl       = card.querySelector('.asinDetail kat-star-rating');
    const productRating = pStarEl?.getAttribute('value') || '';
    const ratingsCount  = pStarEl?.getAttribute('review') || '';
    rows.push([childAsin, createdDate, 'N', reviewer, rating, title, body,
               productRating, ratingsCount, '{domain_code}', '{country}', reviewLink, '', reviewId]);
  }});
  return rows;
}}
"""


SCROLL_JS = """
() => new Promise(resolve => {
  const maxScroll = document.body.scrollHeight * 0.75;
  let pos = 0;
  const tick = () => {
    pos += Math.random() * 180 + 60;
    window.scrollTo(0, Math.min(pos, maxScroll));
    if (pos < maxScroll) setTimeout(tick, Math.random() * 120 + 60);
    else resolve();
  };
  tick();
})
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


async def scrape_domain(domain, page, ctx, prof, asin_filter, out_file=None, append=False):
    """Scrape one domain end-to-end. Returns (total_rows, total_with_imgs).

    out_file : override output path (used by EU group to share one CSV).
    append   : skip header write and load existing rows from out_file first
               (used for EU sub-countries 2-5 so they append to the shared file).
    """
    dc         = _DOMAINS[domain]
    out_file   = out_file or _out_file(domain)
    extract_js = _make_extract_js(domain, dc["country"])
    fetch_js   = _make_batch_fetch_js(dc["review_url"])
    params     = f"?pageSize=50&stars={STAR_FILTER}"
    if "marketplace_id" in dc:
        params += f"&marketplaceId={dc['marketplace_id']}"
    jitter     = prof["fetch_jitter"]
    batch_min  = prof["batch_min"]
    batch_max  = prof["batch_max"]

    print(f"\n{'═'*60}")
    print(f"  Domain : {domain}  ({dc['sc_base']})")
    print(f"  Pages  : {PAGES}  |  Stars: {STAR_FILTER}  |  Output: {out_file}")
    print(f"{'═'*60}")

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

    p = START_PAGE if not append else 1
    while p <= PAGES:
        url = dc["sc_base"] + params + (f"&pageNumber={p}" if p > 1 else "")
        print(f"  Page {p}/{PAGES} …", end=" ", flush=True)
        try:
            await page.goto(url, wait_until="domcontentloaded")
            await page.wait_for_selector('.reviewContainer[data-testid]', timeout=15000)
        except Exception as e:
            print(f"SKIP (timeout/error: {e})")
            p += 1
            continue
        await simulate_reading(page, prof)
        rows = await page.evaluate(extract_js)
        print(f"{len(rows)} reviews  →  flushed to CSV")
        all_rows.extend(rows)
        _csv_append_rows(out_file, rows)   # ← incremental write after every page
        p += 1
        if p <= PAGES:
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

    # ── Step 3: image enrichment ──────────────────────────────────────────
    print(f"  Switching to {dc['amazon_home']} for image fetching …")
    await page.goto(dc["amazon_home"], wait_until="domcontentloaded")
    await asyncio.sleep(random.uniform(1.5, 3.0))
    page = ctx.pages[0] if ctx.pages else page

    review_ids      = [row[IDX['Review ID']] for row in all_rows]
    id_to_row       = {row[IDX['Review ID']]: row for row in all_rows}
    total_with_imgs = 0
    i               = 0

    print(f"  Image fetch     : batches {batch_min}–{batch_max}, jitter {jitter[0]}–{jitter[1]}ms")
    while i < len(review_ids):
        batch   = review_ids[i:i + random.randint(batch_min, batch_max)]
        results = await page.evaluate(fetch_js, [batch, jitter[0], jitter[1]])
        for rid, imgs in results.items():
            if imgs:
                id_to_row[rid][IDX['사진 유무']] = 'Y'
                id_to_row[rid][IDX['Image URL']] = '|'.join(imgs)
                total_with_imgs += 1
        i   += len(batch)
        done = min(i, len(review_ids))
        print(f"    {done}/{len(review_ids)}  ({total_with_imgs} with images)")
        if done < len(review_ids):
            await asyncio.sleep(random.uniform(*prof["batch_delay"]))

    # ── Step 4: ASIN filter ───────────────────────────────────────────────
    if asin_filter:
        before   = len(all_rows)
        all_rows = [r for r in all_rows if r[IDX['ASIN']] in asin_filter]
        print(f"  ASIN filter     : kept {len(all_rows)}/{before}")

    # ── Step 5: column filter ─────────────────────────────────────────────
    if HEADERS_TO_INCLUDE:
        keep_idx    = [IDX[h] for h in HEADERS_TO_INCLUDE if h in IDX]
        out_headers = [ALL_HEADERS[k] for k in keep_idx]
        out_rows    = [[row[k] for k in keep_idx] for row in all_rows]
    else:
        out_headers = ALL_HEADERS
        out_rows    = all_rows

    # ── Step 6: final rewrite with image data ─────────────────────────────
    # In append mode (EU group), merge with enriched rows already in the file
    # from previous sub-countries before rewriting.
    if append and os.path.exists(out_file):
        with open(out_file, encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            next(reader, None)  # skip header
            prev_rows = list(reader)
        out_rows = prev_rows + out_rows
    _csv_rewrite(out_file, out_headers, out_rows)
    print(f"\n  ✓ {domain} done — {total_with_imgs}/{len(all_rows)} with images → {out_file}")

    return len(all_rows), total_with_imgs


async def main():
    if DETECTION_AVOIDANCE not in _PROFILES:
        raise ValueError("DETECTION_AVOIDANCE must be LOW, MEDIUM, or HIGH.")
    unknown = [d for d in DOMAINS if d not in _DOMAINS and d != "EU"]
    if unknown:
        raise ValueError(f"Unknown domain(s): {unknown}. Choose from: EU | {list(_DOMAINS)}")

    prof = _PROFILES[DETECTION_AVOIDANCE]

    asin_filter = None
    if ASIN_FILTER_FILE:
        with open(os.path.expanduser(ASIN_FILTER_FILE), encoding='utf-8') as af:
            asin_filter = {line.strip() for line in af if line.strip()}
        print(f"ASIN filter loaded: {len(asin_filter)} ASINs from {ASIN_FILTER_FILE}")

    print(f"Domains     : {DOMAINS}")
    print(f"Pages/domain: {PAGES}  |  Stars: {STAR_FILTER}  |  Avoidance: {DETECTION_AVOIDANCE}")
    print(f"Headless    : {HEADLESS}")

    async with async_playwright() as pw:
        if HEADLESS:
            ctx     = await pw.chromium.launch_persistent_context(
                CHROME_USER_DATA, channel="chrome", headless=True)
            browser = None
        else:
            browser = await pw.chromium.connect_over_cdp("http://localhost:9222")
            ctx     = browser.contexts[0]

        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        summary = []
        for domain in DOMAINS:
            if domain == "EU":
                # Scrape all EU sub-countries into one combined EU_*.csv
                eu_file = os.path.join(OUT_DIR, "EU_seller_central_reviews.csv")
                eu_rows, eu_imgs = 0, 0
                try:
                    for i, sub in enumerate(EU_COUNTRIES):
                        n_rows, n_imgs = await scrape_domain(
                            sub, page, ctx, prof, asin_filter,
                            out_file=eu_file, append=(i > 0)
                        )
                        eu_rows += n_rows
                        eu_imgs += n_imgs
                    summary.append(("EU", eu_rows, eu_imgs, "OK"))
                except Exception as e:
                    print(f"\n  ✗ EU failed: {e}")
                    summary.append(("EU", eu_rows, eu_imgs, f"FAILED: {e}"))
            else:
                try:
                    n_rows, n_imgs = await scrape_domain(domain, page, ctx, prof, asin_filter)
                    summary.append((domain, n_rows, n_imgs, "OK"))
                except Exception as e:
                    print(f"\n  ✗ {domain} failed: {e}")
                    summary.append((domain, 0, 0, f"FAILED: {e}"))

        if HEADLESS:
            await ctx.close()
        else:
            await browser.close()

    print(f"\n{'═'*60}")
    print("  SUMMARY")
    print(f"{'═'*60}")
    for domain, n_rows, n_imgs, status in summary:
        print(f"  {domain:4s}  {n_rows:>5} reviews  {n_imgs:>4} with images  [{status}]")
    print(f"{'═'*60}")


if __name__ == '__main__':
    asyncio.run(main())

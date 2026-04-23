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

DOMAIN = "US"
# Marketplace to scrape.
# Fully verified : "US"
# Configured, verify SC URL before use: "UK" | "DE" | "FR" | "IT" | "ES" | "JP" | "IN"

PAGES = 30
# Max pages to scrape (50 reviews per page).
# US full scrape ≈ 49 pages (~2,449 reviews).

STAR_FILTER = "1,2,3"
# Comma-separated star ratings to include.
# Critical reviews only: "1,2,3"   All reviews: "1,2,3,4,5"

OUT_FILE = None
# Output CSV path. None → auto-named ~/Desktop/<DOMAIN>_seller_central_reviews.csv

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

# ═══════════════════════════════════════════════════════════════════════════════
# DOMAIN REGISTRY
# Add new marketplaces here as they are onboarded.
# NOTE: EU domains share sellercentral-europe.amazon.com — verify that the
#       active SC session is switched to the correct marketplace before running.
# ═══════════════════════════════════════════════════════════════════════════════

_DOMAINS = {
    "US": {
        "sc_base":     "https://sellercentral.amazon.com/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.com/",
        "review_url":  "https://www.amazon.com/gp/customer-reviews/",
        "country":     "US",
    },
    "UK": {
        "sc_base":     "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.co.uk/",
        "review_url":  "https://www.amazon.co.uk/gp/customer-reviews/",
        "country":     "UK",
    },
    "DE": {
        "sc_base":     "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.de/",
        "review_url":  "https://www.amazon.de/gp/customer-reviews/",
        "country":     "DE",
    },
    "FR": {
        "sc_base":     "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.fr/",
        "review_url":  "https://www.amazon.fr/gp/customer-reviews/",
        "country":     "FR",
    },
    "IT": {
        "sc_base":     "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.it/",
        "review_url":  "https://www.amazon.it/gp/customer-reviews/",
        "country":     "IT",
    },
    "ES": {
        "sc_base":     "https://sellercentral-europe.amazon.com/brand-customer-reviews/",
        "amazon_home": "https://www.amazon.es/",
        "review_url":  "https://www.amazon.es/gp/customer-reviews/",
        "country":     "ES",
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
        "nav_delay":    (0.5,  1.5),   # pause between SC page navigations (s)
        "read_delay":   (0.3,  0.8),   # pause after page load before extracting (s)
        "batch_delay":  (0.5,  1.5),   # pause between image-fetch batches (s)
        "fetch_jitter": (0,    150),   # per-request stagger within a batch (ms)
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


def _resolve_config():
    if DOMAIN not in _DOMAINS:
        raise ValueError(f"Unknown DOMAIN '{DOMAIN}'. Choose from: {list(_DOMAINS)}")
    if DETECTION_AVOIDANCE not in _PROFILES:
        raise ValueError(f"DETECTION_AVOIDANCE must be LOW, MEDIUM, or HIGH.")
    dc   = _DOMAINS[DOMAIN]
    prof = _PROFILES[DETECTION_AVOIDANCE]
    out  = OUT_FILE or os.path.expanduser(f"~/Desktop/{DOMAIN}_seller_central_reviews.csv")
    return dc, prof, out


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


async def main():
    dc, prof, out_file = _resolve_config()

    print(f"  Domain:     {DOMAIN}")
    print(f"  Pages:      {PAGES} (up to {PAGES * 50} reviews)")
    print(f"  Stars:      {STAR_FILTER}")
    print(f"  Avoidance:  {DETECTION_AVOIDANCE}")
    print(f"  Output:     {out_file}")
    print()

    extract_js     = _make_extract_js(DOMAIN, dc["country"])
    batch_fetch_js = _make_batch_fetch_js(dc["review_url"])
    params         = f"?pageSize=50&stars={STAR_FILTER}"

    async with async_playwright() as pw:
        browser = await pw.chromium.connect_over_cdp("http://localhost:9222")
        ctx  = browser.contexts[0]
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        # ── Step 1: scrape Seller Central listing pages ───────────────────
        all_rows = []
        p = 1
        while p <= PAGES:
            url = dc["sc_base"] + params + (f"&pageNumber={p}" if p > 1 else "")
            print(f"  Page {p}/{PAGES} …", end=" ", flush=True)
            await page.goto(url, wait_until="domcontentloaded")
            await page.wait_for_selector('.reviewContainer[data-testid]', timeout=15000)
            await simulate_reading(page, prof)
            rows = await page.evaluate(extract_js)
            print(f"{len(rows)} reviews")
            all_rows.extend(rows)
            p += 1
            if p <= PAGES:
                await asyncio.sleep(random.uniform(*prof["nav_delay"]))

        print(f"\n  Total reviews collected: {len(all_rows)}")

        # ── Step 2: deduplicate by Review ID ─────────────────────────────
        seen_ids, unique_rows = set(), []
        for row in all_rows:
            rid = row[IDX['Review ID']]
            if rid and rid not in seen_ids:
                seen_ids.add(rid)
                unique_rows.append(row)
        dupes = len(all_rows) - len(unique_rows)
        if dupes:
            print(f"  Removed {dupes} duplicate review(s) → {len(unique_rows)} unique")
        all_rows = unique_rows

        # ── Step 3: reviewer image enrichment ────────────────────────────
        print(f"  Switching to {dc['amazon_home']} for image fetching …")
        await page.goto(dc["amazon_home"], wait_until="domcontentloaded")
        await asyncio.sleep(random.uniform(1.5, 3.0))
        page = ctx.pages[0]

        review_ids      = [row[IDX['Review ID']] for row in all_rows]
        id_to_row       = {row[IDX['Review ID']]: row for row in all_rows}
        total_with_imgs = 0
        i               = 0
        jitter          = prof["fetch_jitter"]
        batch_min       = prof["batch_min"]
        batch_max       = prof["batch_max"]

        print(f"  Fetching reviewer images (batches of {batch_min}–{batch_max}, "
              f"jitter {jitter[0]}–{jitter[1]}ms) …")

        while i < len(review_ids):
            batch   = review_ids[i:i + random.randint(batch_min, batch_max)]
            results = await page.evaluate(batch_fetch_js, [batch, jitter[0], jitter[1]])

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

        # ── Step 4: apply ASIN filter if requested ───────────────────────
        if ASIN_FILTER_FILE:
            with open(os.path.expanduser(ASIN_FILTER_FILE), encoding='utf-8') as af:
                target_asins = {line.strip() for line in af if line.strip()}
            before = len(all_rows)
            all_rows = [r for r in all_rows if r[IDX['ASIN']] in target_asins]
            print(f"  ASIN filter: kept {len(all_rows)}/{before} reviews "
                  f"({len(target_asins)} target ASINs)")

        # ── Step 5: apply column filter if requested ──────────────────────
        if HEADERS_TO_INCLUDE:
            keep_idx    = [IDX[h] for h in HEADERS_TO_INCLUDE if h in IDX]
            out_headers = [ALL_HEADERS[i] for i in keep_idx]
            out_rows    = [[row[i] for i in keep_idx] for row in all_rows]
        else:
            out_headers = ALL_HEADERS
            out_rows    = all_rows

        # ── Step 6: write CSV ─────────────────────────────────────────────
        with open(out_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(out_headers)
            writer.writerows(out_rows)

        print(f"\nDone. {total_with_imgs}/{len(all_rows)} reviews have images.")
        print(f"Saved → {out_file}")
        await browser.close()


if __name__ == '__main__':
    asyncio.run(main())

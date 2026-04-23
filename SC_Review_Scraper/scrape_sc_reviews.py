#!/usr/bin/env python3
"""Scrape 1-3 star reviews from Amazon Seller Central US.
Anti-bot measures applied throughout:
  - Random delays between every page navigation
  - Human-like scroll simulation before data extraction
  - Staggered concurrent image fetches with per-request jitter
  - Realistic browser headers on every amazon.com request
  - Random batch sizing so traffic pattern is never uniform
"""

import asyncio, csv, random, re
from playwright.async_api import async_playwright

BASE_URL = "https://sellercentral.amazon.com/brand-customer-reviews/"
PARAMS   = "?pageSize=50&stars=1,2,3"
PAGES    = 30        # set to ~49 for full scrape of all 2,449 reviews
OUT_FILE = "/Users/kevinkim/Desktop/US_seller_central_reviews.csv"

# ── timing knobs ──────────────────────────────────────────────────────────────
NAV_DELAY    = (2.0, 5.0)   # random pause (s) between Seller Central page loads
READ_DELAY   = (1.0, 2.5)   # random pause after page load before extracting
BATCH_DELAY  = (2.0, 4.5)   # random pause between image-fetch batches
FETCH_JITTER = (0,   600)   # per-request stagger within a batch (ms)
BATCH_MIN    = 15            # image-fetch batch size varies between MIN and MAX
BATCH_MAX    = 22

HEADERS = ['ASIN','Created 날짜','사진 유무','Reviewer','Review Ratings','Review Title',
           '본문','Product Rating','Ratings Count','Domain Code','국가',
           'Review Link','Image URL','Review ID']
IDX = {h: i for i, h in enumerate(HEADERS)}

# ── Seller Central extraction ─────────────────────────────────────────────────
EXTRACT_JS = """
() => {
  const cards = document.querySelectorAll('.reviewContainer[data-testid]');
  const rows = [];
  cards.forEach(card => {
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
    card.querySelectorAll('.css-yyccc7').forEach(r => {
      const label = r.querySelector('.css-1ggdaz4')?.textContent?.trim();
      const val   = r.querySelectorAll('div')[1]?.textContent?.trim();
      if (label === 'Child ASIN') childAsin = val;
    });
    const pStarEl       = card.querySelector('.asinDetail kat-star-rating');
    const productRating = pStarEl?.getAttribute('value') || '';
    const ratingsCount  = pStarEl?.getAttribute('review') || '';
    rows.push([childAsin, createdDate, 'N', reviewer, rating, title, body,
               productRating, ratingsCount, 'US', 'US', reviewLink, '', reviewId]);
  });
  return rows;
}
"""

# ── human-like scroll simulation ──────────────────────────────────────────────
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

# ── reviewer image batch fetch with staggered jitter ─────────────────────────
# Each request fires after a random per-request delay so bursts are never uniform.
BATCH_FETCH_JS = """
async (args) => {
  const reviewIds = args[0];
  const jitterMin = args[1];
  const jitterMax = args[2];
  const results = {};
  await Promise.all(reviewIds.map(async (id, idx) => {
    await new Promise(r => setTimeout(r, idx * (Math.random() * (jitterMax - jitterMin) + jitterMin)));
    try {
      const resp = await fetch('https://www.amazon.com/gp/customer-reviews/' + id, {
        credentials: 'include',
        headers: {
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
          'Accept-Language': 'en-US,en;q=0.9',
          'Accept-Encoding': 'gzip, deflate, br',
          'Referer': 'https://www.amazon.com/',
          'Sec-Fetch-Dest': 'document',
          'Sec-Fetch-Mode': 'navigate',
          'Sec-Fetch-Site': 'same-origin',
          'Upgrade-Insecure-Requests': '1'
        }
      });
      const html = await resp.text();
      const doc  = new DOMParser().parseFromString(html, 'text/html');
      const tiles = doc.querySelectorAll('[data-hook="review-image-tile"]');
      results[id] = [...tiles].map(el => {
        const src = el.getAttribute('src') || el.querySelector('img')?.getAttribute('src') || '';
        return src.replace(/\\._[A-Z0-9_,]+_\\./, '.');
      }).filter(Boolean);
    } catch(e) {
      results[id] = [];
    }
  }));
  return results;
}
"""


def rand_batch_size():
    return random.randint(BATCH_MIN, BATCH_MAX)


async def simulate_reading(page):
    """Scroll through the page like a human before extracting data."""
    try:
        await page.evaluate(SCROLL_JS)
    except Exception:
        pass
    await asyncio.sleep(random.uniform(*READ_DELAY))


async def main():
    async with async_playwright() as pw:
        browser = await pw.chromium.connect_over_cdp("http://localhost:9222")
        ctx  = browser.contexts[0]
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        # ── Step 1: scrape Seller Central listing pages ───────────────────
        all_rows = []
        p = 1
        while p <= PAGES:
            url = f"{BASE_URL}{PARAMS}" + (f"&pageNumber={p}" if p > 1 else "")
            print(f"  Page {p}/{PAGES} …", end=" ", flush=True)
            await page.goto(url, wait_until="domcontentloaded")
            await page.wait_for_selector('.reviewContainer[data-testid]', timeout=15000)

            # simulate human reading before scraping
            await simulate_reading(page)

            rows = await page.evaluate(EXTRACT_JS)
            print(f"{len(rows)} reviews")
            all_rows.extend(rows)
            p += 1

            if p <= PAGES:
                await asyncio.sleep(random.uniform(*NAV_DELAY))

        print(f"\n  Total reviews collected: {len(all_rows)}")

        # ── Step 2: reviewer image enrichment ────────────────────────────
        # Deduplicate all_rows by Review ID (keep first occurrence) so that
        # the image fetch counter and CSV row count always agree.
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

        # Navigate to amazon.com — image fetches must be same-origin to
        # include session cookies without triggering CORS/CAPTCHA.
        print("  Switching to amazon.com for image fetching …")
        await page.goto("https://www.amazon.com/", wait_until="domcontentloaded")
        await asyncio.sleep(random.uniform(1.5, 3.0))
        # Re-acquire page reference after navigation (CDP context may shift)
        page = ctx.pages[0]

        review_ids     = [row[IDX['Review ID']] for row in all_rows]
        id_to_row      = {row[IDX['Review ID']]: row for row in all_rows}
        total_with_imgs = 0
        i = 0

        print(f"  Fetching reviewer images (batches of {BATCH_MIN}–{BATCH_MAX}, "
              f"jitter {FETCH_JITTER[0]}–{FETCH_JITTER[1]}ms) …")

        while i < len(review_ids):
            batch_size = rand_batch_size()
            batch      = review_ids[i:i + batch_size]
            results    = await page.evaluate(BATCH_FETCH_JS, [batch, FETCH_JITTER[0], FETCH_JITTER[1]])

            for rid, imgs in results.items():
                if imgs:
                    id_to_row[rid][IDX['사진 유무']] = 'Y'
                    id_to_row[rid][IDX['Image URL']] = '|'.join(imgs)
                    total_with_imgs += 1

            i   += batch_size
            done = min(i, len(review_ids))
            print(f"    {done}/{len(review_ids)}  ({total_with_imgs} with images)")

            if done < len(review_ids):
                await asyncio.sleep(random.uniform(*BATCH_DELAY))

        # ── Step 3: write CSV ─────────────────────────────────────────────
        with open(OUT_FILE, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(HEADERS)
            writer.writerows(all_rows)

        print(f"\nDone. {total_with_imgs}/{len(all_rows)} reviews have images.")
        print(f"Saved → {OUT_FILE}")
        await browser.close()


if __name__ == '__main__':
    asyncio.run(main())

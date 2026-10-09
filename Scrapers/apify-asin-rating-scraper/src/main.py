"""Amazon ASIN rating / review-count scraper (Apify Actor)."""
import asyncio
import re
from datetime import datetime, timezone

from apify import Actor
from curl_cffi.requests import AsyncSession
from parsel import Selector

ASIN_RE = re.compile(r'(?:/dp/|/gp/product/|^)([A-Z0-9]{10})(?:[/?#]|$)', re.I)
MAX_ATTEMPTS = 6


def to_asin(s: str) -> str | None:
    m = ASIN_RE.search(s.strip())
    return m.group(1).upper() if m else None


def num(text: str | None) -> float | None:
    if not text:
        return None
    m = re.search(r'\d[\d.,\s  ]*', text)
    return float(re.sub(r'[^\d]', '', m.group(0))) if m else None


def rating(text: str | None) -> float | None:
    m = re.search(r'(\d+[.,]\d+|\d+)', text or '')
    return float(m.group(1).replace(',', '.')) if m else None


def parse(html: str, asin: str, domain: str) -> dict | None:
    sel = Selector(html)
    if sel.css('form[action*="validateCaptcha"]') or 'api-services-support@amazon.com' in html:
        return None  # blocked -> retry
    title = (sel.css('#productTitle::text').get() or '').strip()
    if not title and not sel.css('#dp, #ppd'):
        if sel.css('img[alt*="Dogs of Amazon"], #g'):  # 404 page
            return {'asin': asin, 'found': False}
        return None
    avg = rating(sel.css('#acrPopover::attr(title)').get()
                 or sel.css('#acrPopover span.a-icon-alt::text').get()
                 or sel.css('[data-hook="rating-out-of-text"]::text').get())
    count = num(sel.css('#acrCustomerReviewText::text').get()
                or sel.css('[data-hook="total-review-count"]::text').get())
    hist = {}
    for row in sel.css('#histogramTable li, #histogramTable tr, ul#histogramTable a'):
        label = ' '.join(row.css('::attr(aria-label)').getall() + row.css('*::text').getall())
        m = re.search(r'(\d)\s*(?:star|Stern|étoile|stell|estrella|つ星|ster)', label, re.I) or re.search(r'(\d)', label)
        p = re.search(r'(\d{1,3})\s*%', label)
        if m and p:
            hist.setdefault(f'star{m.group(1)}Pct', int(p.group(1)))
    return {
        'asin': asin,
        'found': True,
        'url': f'https://www.amazon.{domain}/dp/{asin}',
        'title': title or None,
        'rating': avg,
        'ratingsCount': int(count) if count is not None else 0 if avg is None else None,
        **{k: hist.get(k) for k in ('star5Pct', 'star4Pct', 'star3Pct', 'star2Pct', 'star1Pct')},
        'scrapedAt': datetime.now(timezone.utc).isoformat(),
    }


async def main() -> None:
    async with Actor:
        inp = await Actor.get_input() or {}
        domain = inp.get('marketplace', 'com')
        asins = list(dict.fromkeys(a for a in (to_asin(x) for x in inp.get('asins', [])) if a))
        if not asins:
            await Actor.fail(status_message='No valid ASINs in input')
            return
        proxy = await Actor.create_proxy_configuration(actor_proxy_input=inp.get('proxyConfiguration'))
        sem = asyncio.Semaphore(inp.get('maxConcurrency', 5))
        done = 0

        async def scrape(asin: str) -> None:
            nonlocal done
            async with sem:
                result = None
                for attempt in range(MAX_ATTEMPTS):
                    proxy_url = await proxy.new_url(session_id=f'{asin}_{attempt}') if proxy else None
                    try:
                        async with AsyncSession(impersonate='chrome', proxy=proxy_url, timeout=40) as s:
                            r = await s.get(f'https://www.amazon.{domain}/dp/{asin}',
                                            headers={'Accept-Language': 'en-US,en;q=0.9'})
                        if r.status_code == 404:
                            result = {'asin': asin, 'found': False}
                            break
                        if r.status_code == 200:
                            result = parse(r.text, asin, domain)
                            if result:
                                break
                        Actor.log.info(f'{asin}: attempt {attempt + 1} blocked/status {r.status_code}')
                    except Exception as e:  # network / proxy errors -> retry
                        Actor.log.info(f'{asin}: attempt {attempt + 1} error {e}')
                if not result:
                    result = {'asin': asin, 'found': None, 'error': f'Blocked after {MAX_ATTEMPTS} attempts'}
                result['marketplace'] = f'amazon.{domain}'
                await Actor.push_data(result)
                if result.get('found'):
                    try:
                        await Actor.charge('asin-scraped')
                    except Exception:
                        pass  # not a pay-per-event run
                done += 1
                await Actor.set_status_message(f'{done}/{len(asins)} ASINs')

        await asyncio.gather(*(scrape(a) for a in asins))

"""Amazon ASIN rating / review-count scraper (Apify Actor)."""
import asyncio
import html as htmllib
import json
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


SUMMARY_RE = re.compile(r'"customerReviewSummary":(\{"rating":[^{}]*"histogram":\{[^{}]*\}\})')


def review_summary(html: str, asin: str) -> dict | None:
    """Amazon embeds {rating, count, asin, histogram} JSON for the product (and its variants)."""
    text = htmllib.unescape(html).replace('\\"', '"')
    for m in SUMMARY_RE.finditer(text):
        try:
            d = json.loads(m.group(1))
        except ValueError:
            continue
        if d.get('asin') == asin:
            return d
    return None


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
    summary = review_summary(html, asin)
    if summary:
        avg = avg if avg is not None else summary.get('rating')
        count = count if count is not None else summary.get('count')
        for word, n in (('five', 5), ('four', 4), ('three', 3), ('two', 2), ('one', 1)):
            hist[f'star{n}Pct'] = (summary.get('histogram') or {}).get(f'{word}Star')
    for row in sel.css('#histogramTable li, #histogramTable tr, ul#histogramTable a'):
        label = ' '.join(row.css('::attr(aria-label)').getall() + row.css('*::text').getall())
        m = re.search(r'(\d)\s*(?:star|Stern|étoile|stell|estrella|つ星|ster)', label, re.I)
        p = re.search(r'(\d{1,3})\s*%', label)
        if m and p and hist.get(f'star{m.group(1)}Pct') is None:
            hist[f'star{m.group(1)}Pct'] = int(p.group(1))
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


async def fetch_dp(s: AsyncSession, asin: str, domain: str):
    """GET the product page, passing Amazon's 'Continue shopping' interstitial if served."""
    base = f'https://www.amazon.{domain}'
    headers = {'Accept-Language': 'en-US,en;q=0.9'}
    r = await s.get(f'{base}/dp/{asin}', headers=headers)
    for _ in range(2):
        form = Selector(r.text).css('form[action*="validateCaptcha"]')
        if r.status_code != 200 or not form or form.css('img'):  # image captcha -> new IP
            break
        params = {i.attrib['name']: i.attrib.get('value', '') for i in form.css('input[name]')}
        await s.get(base + form.attrib['action'], params=params, headers=headers)
        r = await s.get(f'{base}/dp/{asin}', headers=headers)
    return r


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
                            r = await fetch_dp(s, asin, domain)
                        if r.status_code == 404:
                            result = {'asin': asin, 'found': False}
                            break
                        if r.status_code == 200:
                            result = parse(r.text, asin, domain)
                            if inp.get('debugHtml'):
                                await Actor.set_value(f'html_{asin}_{attempt}', r.text, content_type='text/html')
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

"""Amazon ASIN rating / review-count scraper (Apify Actor)."""
import asyncio
import html
import os
import random
import re
from datetime import datetime, timezone

from apify import Actor
from curl_cffi.requests import AsyncSession
from parsel import Selector

ASIN_RE = re.compile(r'(?:/dp/|/gp/product/|^)([A-Z0-9]{10})(?:[/?#]|$)', re.I)
MAX_ATTEMPTS = 6
DATACENTER_ATTEMPTS = 2
MAX_CONCURRENCY = 50
MB_PER_SESSION = 12  # measured ~6 MB per session with byte-level parsing (21 sessions peaked at 173 MB); 2x headroom
SESSION_MAX_PAGES = 40  # rotate IP/cookies after this many pages, like a real visitor leaving


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


TITLE_RE = re.compile(rb'id="productTitle"[^>]*>\s*([^<]+?)\s*<')
POPOVER_RE = re.compile(rb'id="acrPopover"[^>]*?title="([^"]+)"')
STARS_ALT_RE = re.compile(rb'id="acrPopover".{0,600}?class="a-icon-alt">([^<]+)<', re.S)
COUNT_RE = re.compile(rb'id="acrCustomerReviewText"[^>]*>([^<]+)<')


def txt(m) -> str:
    return html.unescape(m.group(1).decode('utf-8', 'replace')).strip() if m else ''


def parse(page: bytes, asin: str, domain: str) -> dict | None:
    """Regex on the raw bytes — no decoding of the 2-3 MB page and no DOM tree, so RAM/CPU per page stay small."""
    if b'validateCaptcha' in page or b'api-services-support@amazon.com' in page:
        return None  # blocked -> retry
    title = txt(TITLE_RE.search(page))
    if not title:
        return None  # degraded/partial page (Amazon serves these intermittently) -> retry on a fresh session
    m = POPOVER_RE.search(page) or STARS_ALT_RE.search(page)
    avg = rating(txt(m)) if m else None
    c = COUNT_RE.search(page)
    count = num(txt(c)) if c else None
    return {
        'asin': asin,
        'found': True,
        'url': f'https://www.amazon.{domain}/dp/{asin}',
        'title': title or None,
        'rating': avg,
        'ratingsCount': int(count) if count is not None else 0 if avg is None else None,
        'scrapedAt': datetime.now(timezone.utc).isoformat(),
    }


HEADERS = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
}
# Real-browser TLS/HTTP2 fingerprints to rotate between sessions.
FINGERPRINTS = ['chrome131', 'chrome124', 'chrome123', 'chrome120', 'edge101', 'safari17_0']


async def fetch_dp(s: AsyncSession, asin: str, domain: str):
    """GET the product page, passing Amazon's 'Continue shopping' interstitial if served."""
    base = f'https://www.amazon.{domain}'
    r = await s.get(f'{base}/dp/{asin}', headers=HEADERS)
    for _ in range(2):
        if r.status_code != 200 or b'validateCaptcha' not in r.content:
            break
        form = Selector(r.text).css('form[action*="validateCaptcha"]')  # small interstitial page only
        if not form or form.css('img'):  # image captcha -> new IP
            break
        params = {i.attrib['name']: i.attrib.get('value', '') for i in form.css('input[name]')}
        await s.get(base + form.attrib['action'], params=params, headers=HEADERS)
        r = await s.get(f'{base}/dp/{asin}', headers=HEADERS)
    return r


class Worker:
    """One simulated browser: a sticky proxy IP + cookie jar + fingerprint, reused across ASINs.

    Reusing the session keeps the cookies Amazon sets after the interstitial, so later pages
    come straight through (fewer requests, fewer blocks). Retired and replaced when blocked.
    """

    def __init__(self, wid: int, proxy, fallback):
        self.wid, self.proxy, self.fallback, self.gen = wid, proxy, fallback, 0
        self.sessions: dict[bool, AsyncSession] = {}
        self.served = 0

    async def session(self, residential: bool) -> AsyncSession:
        if residential not in self.sessions:
            cfg = self.fallback if residential else self.proxy
            url = await cfg.new_url(session_id=f'w{self.wid}g{self.gen}{"r" if residential else "d"}') if cfg else None
            fp = FINGERPRINTS[(self.wid + self.gen) % len(FINGERPRINTS)]
            self.sessions[residential] = AsyncSession(impersonate=fp, proxy=url, timeout=25)
        return self.sessions[residential]

    async def retire(self, residential: bool) -> None:
        s = self.sessions.pop(residential, None)
        if s:
            await s.close()
        self.gen += 1
        self.served = 0

    async def close(self) -> None:
        for s in self.sessions.values():
            await s.close()


async def main() -> None:
    async with Actor:
        inp = await Actor.get_input() or {}
        domain = inp.get('marketplace', 'com')
        asins = list(dict.fromkeys(a for a in (to_asin(x) for x in inp.get('asins', [])) if a))
        if not asins:
            await Actor.fail(status_message='No valid ASINs in input')
            return
        proxy_input = inp.get('proxyConfiguration') or {'useApifyProxy': True}
        proxy = await Actor.create_proxy_configuration(actor_proxy_input=proxy_input)
        # Cheap datacenter IPs first; escalate to residential only for ASINs that keep getting blocked.
        fallback = None
        if proxy_input.get('useApifyProxy') and not proxy_input.get('apifyProxyGroups'):
            fallback = await Actor.create_proxy_configuration(groups=['RESIDENTIAL'])
        memory_mb = int(os.environ.get('ACTOR_MEMORY_MBYTES') or os.environ.get('APIFY_MEMORY_MBYTES') or 1024)
        safe = max(4, memory_mb // MB_PER_SESSION)
        requested = int(inp.get('maxConcurrency', 40))
        concurrency = max(1, min(requested, safe, MAX_CONCURRENCY, len(asins)))
        if requested > safe:
            Actor.log.info(f'maxConcurrency {requested} capped to {safe} for {memory_mb} MB memory — give the run more memory for more speed')
        queue: asyncio.Queue = asyncio.Queue()
        for a in asins:
            queue.put_nowait((a, 0))
        done = 0
        blocks = 0  # recent blocks; workers back off when Amazon pushes back
        lock = asyncio.Lock()

        async def finish(asin: str, result: dict) -> None:
            nonlocal done
            result['marketplace'] = f'amazon.{domain}'
            await Actor.push_data(result)
            if result.get('found'):
                try:
                    await Actor.charge('asin-scraped')
                except Exception:
                    pass  # not a pay-per-event run
            async with lock:
                done += 1
                if done % 25 == 0 or done == len(asins):
                    await Actor.set_status_message(f'{done}/{len(asins)} ASINs')

        async def run_worker(w: Worker) -> None:
            nonlocal blocks
            await asyncio.sleep(random.uniform(0, 2))  # stagger start, no burst of identical requests
            try:
                while True:
                    try:
                        asin, attempt = queue.get_nowait()
                    except asyncio.QueueEmpty:
                        return
                    residential = bool(fallback) and attempt >= DATACENTER_ATTEMPTS
                    result = None
                    try:
                        if w.served >= SESSION_MAX_PAGES:
                            await w.retire(residential)
                        r = await fetch_dp(await w.session(residential), asin, domain)
                        w.served += 1
                        if r.status_code == 404:
                            result = {'asin': asin, 'found': False}
                        elif r.status_code == 200:
                            result = parse(r.content, asin, domain)
                            if inp.get('debugHtml'):
                                await Actor.set_value(f'html_{asin}_{attempt}', r.text, content_type='text/html')
                        if not result:
                            Actor.log.info(f'{asin}: attempt {attempt + 1} blocked/status {r.status_code}')
                    except Exception as e:  # network / proxy errors -> retry on a fresh session
                        Actor.log.info(f'{asin}: attempt {attempt + 1} error {e}')
                    if result:
                        blocks = max(0, blocks - 1)
                        await finish(asin, result)
                    else:
                        blocks += 1
                        await w.retire(residential)
                        if attempt + 1 >= MAX_ATTEMPTS:
                            await finish(asin, {'asin': asin, 'found': None, 'error': f'Blocked after {MAX_ATTEMPTS} attempts'})
                        else:
                            queue.put_nowait((asin, attempt + 1))
                        # adaptive backoff: the more blocks recently, the longer this worker cools down
                        await asyncio.sleep(min(30, 0.5 * 2 ** min(blocks, 6)) * random.uniform(0.5, 1.0))
                        continue
                    await asyncio.sleep(random.uniform(0.2, 0.8))  # human-ish pacing per session
            finally:
                await w.close()

        Actor.log.info(f'{len(asins)} ASINs on amazon.{domain} with {concurrency} parallel sessions ({memory_mb} MB)')
        await asyncio.gather(*(run_worker(Worker(i, proxy, fallback)) for i in range(concurrency)))

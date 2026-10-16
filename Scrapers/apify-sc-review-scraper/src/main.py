"""Amazon Seller Central Brand Customer Reviews scraper (Apify Actor).

Port of the GCX SC_Review_Scraper (scrape_sc_reviews.py) for Apify: instead of logging in with a
password + OTP, the user pastes their own Seller Central login session (cookies) per region.
Output columns (same order as the GCX SC scraper / Caspi fetch, English headers):
  ASIN | Created Date | Has Photo | Reviewer | Review Ratings | Review Title | Review Text | Country |
  Review Link | Image URL | Review ID | Order ID | Product Rating | Ratings Count
"""
import asyncio
import random
import re
from datetime import datetime, timezone
from urllib.parse import parse_qs, urlparse

from apify import Actor
from playwright.async_api import async_playwright

from .sessions import fingerprint, parse_session

# ── Marketplace registry (same as scrape_sc_reviews.py _DOMAINS) ─────────────────────────────
REGIONS = {
    'US': {'host': 'sellercentral.amazon.com'},
    'EU': {'host': 'sellercentral-europe.amazon.com'},
    'JP': {'host': 'sellercentral.amazon.co.jp'},
    'IN': {'host': 'sellercentral.amazon.in'},
}
MARKETS = {
    'US': {'region': 'US', 'amazon': 'www.amazon.com'},
    'DE': {'region': 'EU', 'amazon': 'www.amazon.de', 'display': 'Germany', 'mkid': 'A1PA6795UKMFR9'},
    'IT': {'region': 'EU', 'amazon': 'www.amazon.it', 'display': 'Italy', 'mkid': 'APJ6JRA9NG5V4'},
    'FR': {'region': 'EU', 'amazon': 'www.amazon.fr', 'display': 'France', 'mkid': 'A13V1IB3VIYZZH'},
    'ES': {'region': 'EU', 'amazon': 'www.amazon.es', 'display': 'Spain', 'mkid': 'A1RKKUPIHCS9HS'},
    'UK': {'region': 'EU', 'amazon': 'www.amazon.co.uk', 'display': 'United Kingdom', 'mkid': 'A1F83G8C2ARO7P'},
    'JP': {'region': 'JP', 'amazon': 'www.amazon.co.jp'},
    'IN': {'region': 'IN', 'amazon': 'www.amazon.in'},
}
EU_ORDER = ['DE', 'IT', 'FR', 'ES', 'UK']  # same order as the SC scraper

PROFILES = {  # detection avoidance — identical to scrape_sc_reviews.py _PROFILES
    'LOW': {'nav_delay': (0.5, 1.5), 'read_delay': (0.3, 0.8), 'batch_delay': (0.5, 1.5),
            'fetch_jitter': (0, 150), 'batch_min': 20, 'batch_max': 30, 'scroll': False},
    'MEDIUM': {'nav_delay': (2.0, 5.0), 'read_delay': (1.0, 2.5), 'batch_delay': (2.0, 4.5),
               'fetch_jitter': (0, 600), 'batch_min': 15, 'batch_max': 22, 'scroll': True},
    'HIGH': {'nav_delay': (4.0, 10.0), 'read_delay': (2.0, 5.0), 'batch_delay': (5.0, 12.0),
             'fetch_jitter': (0, 1200), 'batch_min': 8, 'batch_max': 15, 'scroll': True},
}

ALL_HEADERS = ['ASIN', 'Created Date', 'Has Photo', 'Reviewer', 'Review Ratings', 'Review Title', 'Review Text',
               'Product Rating', 'Ratings Count', 'Domain Code', 'Country', 'Review Link', 'Image URL',
               'Review ID', 'Order ID']
OUTPUT_HEADERS = ['ASIN', 'Created Date', 'Has Photo', 'Reviewer', 'Review Ratings', 'Review Title', 'Review Text',
                  'Country', 'Review Link', 'Image URL', 'Review ID', 'Order ID', 'Product Rating', 'Ratings Count']
ENGLISH_KEYS = {'ASIN': 'asin', 'Created Date': 'createdDate', 'Has Photo': 'hasPhoto', 'Reviewer': 'reviewer',
                'Review Ratings': 'reviewRating', 'Review Title': 'reviewTitle', 'Review Text': 'reviewText',
                'Country': 'marketplace', 'Review Link': 'reviewLink', 'Image URL': 'imageUrls',
                'Review ID': 'reviewId', 'Order ID': 'orderId', 'Product Rating': 'productRating',
                'Ratings Count': 'ratingsCount'}
IDX = {h: i for i, h in enumerate(ALL_HEADERS)}
LOGIN_MARKERS = ('/ap/signin', '/ap/mfa', '/ap/cvf', 'signin', '/mfa')

# ── In-page JS, unchanged from scrape_sc_reviews.py ─────────────────────────────────────────


def make_extract_js(domain_code, country):
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
      const prodLink = card.querySelector('kat-link[href*="/dp/"], kat-link[href*="ASIN="]')?.getAttribute('href') || '';
      const asinFromLink = prodLink.match(/\\/dp\\/([A-Z0-9]{{10}})/)?.[1]
                        || new URLSearchParams(prodLink.split('?')[1] || '').get('ASIN') || '';
      childAsin = asinFromLink || _meta['ASIN'] || '';
    }}
    const pStarEl       = card.querySelector('.asinDetail kat-star-rating');
    const productRating = pStarEl?.getAttribute('value') || '';
    const ratingsCount  = pStarEl?.getAttribute('review') || '';
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
  const cap = setTimeout(resolve, 8000);
  const tick = () => {
    pos += Math.random() * 180 + 60;
    window.scrollTo(0, Math.min(pos, maxScroll));
    if (pos < maxScroll) setTimeout(tick, Math.random() * 120 + 60);
    else { clearTimeout(cap); resolve(); }
  };
  tick();
})
"""

CONTENT_READY_JS = """
() => {
    const cards = document.querySelectorAll('.reviewContainer[data-testid]');
    if (!cards.length) return false;
    return [...cards].every(c => (c.querySelector('.css-g7g1lz')?.textContent?.trim() || '').length > 0);
}
"""

ORDER_ID_FETCH_JS = """async (url) => {
    const res = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/json' } });
    const text = await res.text();
    if (!res.ok) return { __error: `HTTP ${res.status}`, __body: text.slice(0, 200) };
    try { return JSON.parse(text); }
    catch (e) { return { __error: `parse: ${e.message}`, __body: text.slice(0, 200) }; }
}"""


def make_image_fetch_js(review_url_base):
    return f"""
async (args) => {{
  const [reviewIds, jitterMin, jitterMax] = args;
  const results = {{}};
  await Promise.all(reviewIds.map(async (id, idx) => {{
    await new Promise(r => setTimeout(r, idx * (Math.random() * (jitterMax - jitterMin) + jitterMin)));
    try {{
      const resp = await fetch('{review_url_base}' + id, {{ credentials: 'include',
        headers: {{ 'Accept': 'text/html,application/xhtml+xml', 'Referer': '{review_url_base}' }} }});
      const doc = new DOMParser().parseFromString(await resp.text(), 'text/html');
      results[id] = [...doc.querySelectorAll('[data-hook="review-image-tile"]')].map(el => {{
        const src = el.getAttribute('src') || el.querySelector('img')?.getAttribute('src') || '';
        return src.replace(/\\._[A-Z0-9_,]+_\\./, '.');
      }}).filter(Boolean);
    }} catch (e) {{ results[id] = []; }}
  }}));
  return results;
}}
"""

# ── Helpers ──────────────────────────────────────────────────────────────────────────────────

DATE_FORMATS = ('%B %d, %Y', '%d %B %Y', '%b %d, %Y', '%d %b %Y', '%Y/%m/%d', '%Y-%m-%d', '%d/%m/%Y')


def normalize_date(s: str) -> str:
    s = (s or '').strip()
    m = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', s)
    if m:
        return f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return s


def stale_ids(rows):
    """Consecutive cards with different IDs but identical Reviewer+Title = recycled DOM slot."""
    out = set()
    for prev, cur in zip(rows, rows[1:]):
        if (cur[IDX['Review ID']] != prev[IDX['Review ID']] and cur[IDX['Reviewer']]
                and cur[IDX['Reviewer']] == prev[IDX['Reviewer']]
                and cur[IDX['Review Title']] and cur[IDX['Review Title']] == prev[IDX['Review Title']]):
            out.add(cur[IDX['Review ID']])
    return out


def on_login_page(url: str) -> bool:
    return any(m in url for m in LOGIN_MARKERS)


class SessionExpired(Exception):
    pass


class Scraper:
    def __init__(self, inp: dict, ctx):
        self.inp, self.ctx = inp, ctx
        self.prof = PROFILES.get(str(inp.get('detectionAvoidance', 'MEDIUM')).upper(), PROFILES['MEDIUM'])
        self.page_size = min((25, 50, 100), key=lambda v: abs(v - int(inp.get('pageSize', 50))))
        self.max_pages = int(inp.get('maxPagesPerMarketplace', 30))
        self.stars = ','.join(str(s) for s in (inp.get('stars') or [1, 2, 3, 4, 5]))
        self.min_date = (inp.get('minReviewDate') or '').strip() or None
        self.asin_filter = {a.strip().upper() for a in (inp.get('asinFilter') or []) if a.strip()}
        self.fetch_order_id = inp.get('fetchOrderId', True)
        self.fetch_images = inp.get('fetchImages', False)
        self.english = inp.get('outputFieldNames', 'title') == 'camel'
        self.health: dict[str, dict] = {}
        self.pushed = 0

    async def sleep(self, key):
        await asyncio.sleep(random.uniform(*self.prof[key]))

    # ── account / marketplace selection ──────────────────────────────────────────────────
    async def open_region(self, page, region: str, url: str):
        await page.goto(url, wait_until='domcontentloaded', timeout=45000)
        if on_login_page(page.url):
            raise SessionExpired(f'{region}: Seller Central redirected to sign-in — the pasted session is expired or incomplete')
        await self.resolve_account_picker(page, region)
        if on_login_page(page.url):
            raise SessionExpired(f'{region}: session expired while selecting the account')

    async def resolve_account_picker(self, page, region):
        """Full-page 'Select an account' picker (multi-account sellers)."""
        try:
            if not await page.get_by_text('Select an account', exact=False).count():
                return
        except Exception:
            return
        label = '.full-page-account-switcher-account-label'
        wanted = (self.inp.get(f'account{region}') or '').strip()
        labels = [t.strip() for t in await page.locator(label).all_inner_texts()]
        if not wanted:
            raise RuntimeError(f"{region}: Seller Central asks to pick an account — set 'Account name ({region})' "
                               f"to one of: {', '.join(dict.fromkeys(labels))[:400]}")
        loc = page.locator(f"{label}:text-is('{wanted}')").first
        if not await loc.count():
            raise RuntimeError(f"{region}: account '{wanted}' not in this session's picker ({', '.join(dict.fromkeys(labels))[:400]})")
        await loc.click()
        await asyncio.sleep(1.2)
        sub = {'US': 'United States', 'EU': 'Germany', 'JP': 'Japan', 'IN': 'India'}.get(region)
        sub_loc = page.locator(f"{label}:text-is('{sub}')").first
        if sub and await sub_loc.count():
            await sub_loc.click()
            await asyncio.sleep(0.8)
        confirm = page.locator("[data-test='confirm-selection']").first
        if await confirm.count():
            await confirm.click()
            await page.wait_for_load_state('domcontentloaded', timeout=20000)
            await asyncio.sleep(1.5)
        Actor.log.info(f'[{region}] account selected: {wanted}')

    async def switch_eu_marketplace(self, page, market: str) -> str:
        """Two-level account switcher on SC Europe → returns mons_sel_* params for that country."""
        m = MARKETS[market]
        await page.goto('https://sellercentral-europe.amazon.com/home', wait_until='domcontentloaded', timeout=45000)
        if on_login_page(page.url):
            raise SessionExpired('EU: Seller Central redirected to sign-in — the pasted EU session is expired')
        await page.wait_for_selector('.dropdown-account-switcher-header', timeout=20000)
        try:
            btn = await page.wait_for_selector('#modal button.modal-close', timeout=2500, state='visible')
            await btn.click()
        except Exception:
            pass
        account = (self.inp.get('accountEU') or '').strip()
        try:
            async with page.expect_navigation(timeout=20000):
                try:
                    await page.evaluate("""async ([account, country]) => {
                        document.querySelector('.dropdown-account-switcher-header').click();
                        const waitFor = (sel, pred, ms=6000) => new Promise((res, rej) => {
                            const t0 = Date.now(); const iv = setInterval(() => {
                                const el = [...document.querySelectorAll(sel)].find(pred);
                                if (el) { clearInterval(iv); res(el); } else if (Date.now()-t0 > ms) { clearInterval(iv); rej(new Error('timeout ' + sel)); }
                            }, 80); });
                        const labels = '.dropdown-account-switcher-list-item-label';
                        const top = account
                            ? await waitFor(labels, e => e.textContent.trim() === account)
                            : await waitFor(labels, e => e.classList.contains('dropdown-account-switcher-list-item-label-expanded'), 1500)
                                .catch(() => waitFor(labels, () => true));
                        if (!top.classList.contains('dropdown-account-switcher-list-item-label-expanded')) top.click();
                        (await waitFor('.dropdown-account-switcher-list-item-indented', e => e.textContent.trim() === country)).click();
                    }""", [account, m['display']])
                except Exception:
                    pass  # navigation destroys the context — expected
        except Exception as e:
            Actor.log.warning(f'[{market}] marketplace switch navigation failed ({e}) — using marketplace id only')
            return f"mons_sel_mkid={m['mkid']}"
        qs = parse_qs(urlparse(page.url).query, keep_blank_values=True)
        mcid, mkid = qs.get('mons_sel_dir_mcid', [''])[0], qs.get('mons_sel_mkid', [''])[0]
        if mcid and mkid:
            return f'mons_sel_dir_mcid={mcid}&mons_sel_mkid={mkid}'
        return '&'.join(f'{k}={v[0]}' for k, v in qs.items() if k.startswith('mons_sel')) or f"mons_sel_mkid={m['mkid']}"

    # ── scraping ─────────────────────────────────────────────────────────────────────────
    async def scrape_market(self, page, market: str, mons: str = '') -> list[list]:
        region = MARKETS[market]['region']
        base = f"https://{REGIONS[region]['host']}/brand-customer-reviews/"
        params = f'?pageSize={self.page_size}&stars={self.stars}' + (f'&{mons}' if mons else '')
        extract = make_extract_js(market, market)
        rows_all, empty, p = [], 0, 1
        while p <= self.max_pages:
            if not self.budget_left():
                Actor.log.warning(f'[{market}] maximum cost per run reached — stopping (raise it in the run options to get more)')
                break
            url = base + params + (f'&pageNumber={p}' if p > 1 else '')
            try:
                await page.goto(url, wait_until='domcontentloaded', timeout=45000)
                if on_login_page(page.url):
                    raise SessionExpired(f'{market}: session expired on page {p}')
                await page.wait_for_selector('.reviewContainer[data-testid]', timeout=20000)
            except SessionExpired:
                raise
            except Exception as e:
                if on_login_page(page.url):
                    raise SessionExpired(f'{market}: session expired on page {p}')
                try:
                    cards = await page.evaluate("document.querySelectorAll('.reviewContainer[data-testid]').length")
                except Exception:
                    cards = -1
                if cards == 0:
                    empty += 1
                    if empty >= 2:
                        Actor.log.info(f'[{market}] reached the end of available reviews at page {p}')
                        break
                else:
                    Actor.log.warning(f'[{market}] page {p} skipped ({str(e)[:120]})')
                p += 1
                continue
            empty = 0
            if p == 1 and 'display' in MARKETS[market]:
                header = await page.evaluate("document.querySelector('.dropdown-account-switcher-header')?.textContent?.trim() || ''")
                if header and MARKETS[market]['display'].lower() not in header.lower():
                    raise RuntimeError(f"{market}: marketplace mismatch — page shows '{header}', not {MARKETS[market]['display']}")
            if self.prof['scroll']:
                try:
                    await page.evaluate(SCROLL_JS)
                except Exception:
                    pass
            await self.sleep('read_delay')
            for fn, t in (('networkidle', 5000), (CONTENT_READY_JS, 8000)):
                try:
                    if fn == 'networkidle':
                        await page.wait_for_load_state(fn, timeout=t)
                    else:
                        await page.wait_for_function(fn, timeout=t)
                except Exception:
                    pass
            rows = await page.evaluate(extract)
            if p == 1 and rows and not any(r[IDX['Reviewer']] for r in rows):
                raise RuntimeError(f'{market}: review cards found but not in English — switch the Seller Central '
                                   f'language to English (top-right language menu) and copy the session again')
            for r in rows:
                r[IDX['Created Date']] = normalize_date(r[IDX['Created Date']])
            bad = stale_ids(rows)
            if bad:  # reload once, as the SC scraper does
                await asyncio.sleep(random.uniform(1.0, 2.5))
                await page.goto(url, wait_until='domcontentloaded', timeout=45000)
                try:
                    await page.wait_for_selector('.reviewContainer[data-testid]', timeout=20000)
                    await page.wait_for_function(CONTENT_READY_JS, timeout=10000)
                except Exception:
                    pass
                rows = await page.evaluate(extract)
                for r in rows:
                    r[IDX['Created Date']] = normalize_date(r[IDX['Created Date']])
                still = stale_ids(rows)
                rows = [r for r in rows if r[IDX['Review ID']] not in still]
            Actor.log.info(f'[{market}] page {p}/{self.max_pages}: {len(rows)} reviews')
            rows_all.extend(rows)
            if self.min_date and rows and all(r[IDX['Created Date']] < self.min_date for r in rows):
                Actor.log.info(f'[{market}] page {p} is older than {self.min_date} — stopping early')
                break
            p += 1
            if p <= self.max_pages:
                await self.sleep('nav_delay')
        seen, unique = set(), []
        for r in rows_all:
            if r[IDX['Review ID']] and r[IDX['Review ID']] not in seen:
                seen.add(r[IDX['Review ID']])
                unique.append(r)
        if self.fetch_order_id and unique:
            await self.add_order_ids(page, region, unique, mons)
        return unique

    async def add_order_ids(self, page, region, rows, mons):
        api = f"https://{REGIONS[region]['host']}/brandcustomerreviews/api/reviews"
        order, pid, total = {}, 0, 1
        want = {r[IDX['Review ID']] for r in rows}
        while pid < total and pid < 400:
            qs = f'?pageId={pid}&pageSize=50&sortByType=REVIEW_CREATED_DATE&isAscending=false&includeDone=false'
            data = await page.evaluate(ORDER_ID_FETCH_JS, api + qs + (f'&{mons}' if mons else ''))
            if not isinstance(data, dict) or data.get('__error'):
                Actor.log.warning(f'[{region}] order ID API: {data.get("__error") if isinstance(data, dict) else data}')
                break
            total = data.get('totalPageCount', 1)
            for rev in data.get('reviews', []):
                if rev.get('reviewId'):
                    order[rev['reviewId']] = rev.get('orderId') or ''
            pid += 1
            if want <= set(order):
                break
            await asyncio.sleep(0.3)
        matched = 0
        for r in rows:
            oid = order.get(r[IDX['Review ID']], '')
            if oid:
                r[IDX['Order ID']] = oid
                matched += 1
        Actor.log.info(f'[{region}] order IDs matched {matched}/{len(rows)}')

    async def add_images(self, page, market, rows):
        host = MARKETS[market]['amazon']
        try:
            await page.goto(f'https://{host}/', wait_until='domcontentloaded', timeout=45000)
        except Exception as e:
            Actor.log.warning(f'[{market}] storefront unreachable for images ({e})')
            return
        js = make_image_fetch_js(f'https://{host}/gp/customer-reviews/')
        ids = [r[IDX['Review ID']] for r in rows]
        by_id = {r[IDX['Review ID']]: r for r in rows}
        found, i = 0, 0
        while i < len(ids):
            batch = ids[i:i + random.randint(self.prof['batch_min'], self.prof['batch_max'])]
            try:
                res = await page.evaluate(js, [batch, *self.prof['fetch_jitter']])
            except Exception:
                res = {}
            for rid, imgs in (res or {}).items():
                if imgs:
                    by_id[rid][IDX['Has Photo']] = 'Y'
                    by_id[rid][IDX['Image URL']] = '|'.join(imgs)
                    found += 1
            i += len(batch)
            if i < len(ids):
                await self.sleep('batch_delay')
        Actor.log.info(f'[{market}] images: {found}/{len(rows)} reviews have photos')

    async def emit(self, rows):
        out = []
        for r in rows:
            if self.asin_filter and r[IDX['ASIN']].upper() not in self.asin_filter:
                continue
            if self.min_date and r[IDX['Created Date']] < self.min_date:
                continue
            item = {h: r[IDX[h]] for h in OUTPUT_HEADERS}
            out.append({ENGLISH_KEYS[k]: v for k, v in item.items()} if self.english else item)
        if out:
            # Pay-per-event: one 'review' event per row; the SDK never charges past the user's max charge.
            await Actor.push_data(out, charged_event_name='review')
            self.pushed += len(out)
        return len(out)

    def budget_left(self) -> bool:
        try:
            return not Actor.get_charging_manager().is_event_charge_limit_reached('review')
        except Exception:
            return True

    async def run_region(self, region: str, markets: list[str]):
        page = await self.ctx.new_page()
        host = REGIONS[region]['host']
        try:
            await self.open_region(page, region, f'https://{host}/brand-customer-reviews/')
            for market in markets:
                try:
                    mons = await self.switch_eu_marketplace(page, market) if region == 'EU' else ''
                    rows = await self.scrape_market(page, market, mons)
                    if self.fetch_images and rows:
                        await self.add_images(page, market, rows)
                    n = await self.emit(rows)
                    self.health[market] = {'status': 'OK', 'reviews': n}
                    await Actor.set_status_message(f'{market}: {n} reviews · total {self.pushed}')
                except SessionExpired as e:
                    raise e
                except Exception as e:
                    Actor.log.exception(f'[{market}] failed')
                    self.health[market] = {'status': 'ERROR', 'message': str(e)[:500]}
        except SessionExpired as e:
            Actor.log.error(str(e))
            for market in markets:
                self.health.setdefault(market, {'status': 'SESSION_EXPIRED',
                                                'message': f'{e}. Copy a fresh session for {region} and run again.'})
        except Exception as e:
            Actor.log.exception(f'[{region}] failed')
            for market in markets:
                self.health.setdefault(market, {'status': 'ERROR', 'message': str(e)[:500]})
        finally:
            await page.close()


async def main() -> None:
    async with Actor:
        inp = await Actor.get_input() or {}
        markets = [m.upper() for m in (inp.get('marketplaces') or ['US'])]
        unknown = [m for m in markets if m not in MARKETS]
        if unknown:
            await Actor.exit(status_message=f'⚠️ Unknown marketplace(s): {unknown}. Use {list(MARKETS)}')
            return
        by_region: dict[str, list[str]] = {}
        for m in (EU_ORDER + [x for x in MARKETS if x not in EU_ORDER]):
            if m in markets:
                by_region.setdefault(MARKETS[m]['region'], []).append(m)

        cookies, missing = [], []
        for region in by_region:
            raw = inp.get(f'session{region}') or ''
            try:
                parsed = parse_session(raw, REGIONS[region]['host'])
            except Exception as e:
                await Actor.exit(status_message=f'⚠️ {region} session could not be read ({e}). Paste the Cookie-Editor JSON export or the Cookie header.')
                return
            if not parsed:
                missing.append(region)
                continue
            Actor.log.info(f'[{region}] session loaded: {fingerprint(parsed)}')
            cookies += parsed
        if inp.get('fetchImages') and inp.get('storefrontSession'):
            store = parse_session(inp['storefrontSession'], 'www.amazon.com')
            Actor.log.info(f'[storefront] session loaded: {fingerprint(store)}')
            cookies += store
        if not cookies:
            # Nothing to authenticate with (e.g. a first look at the actor, or Apify's automated
            # daily test run on the prefilled input) — explain the setup instead of failing.
            msg = ('No Seller Central session pasted. Paste a session for each region you selected '
                   '(see "How to copy your login session" in the README), then start again.')
            Actor.log.warning(msg)
            await Actor.set_value('SESSION_HEALTH', {'finishedAt': datetime.now(timezone.utc).isoformat(), 'reviews': 0,
                                                     'marketplaces': {m: {'status': 'NO_SESSION'} for m in markets}})
            await Actor.exit(status_message=msg)
            return
        for region in missing:
            Actor.log.error(f'No session pasted for {region} — skipping {by_region.pop(region)}')

        Actor.log.info('Reminder: Seller Central sessions expire. Refresh (re-copy) them before each scheduled '
                       'run or at least every few days so runs stay healthy.')
        health: dict = {m: {'status': 'NO_SESSION', 'message': 'No session pasted for this region'}
                        for r in missing for m in [x for x in markets if MARKETS[x]['region'] == r]}
        async with async_playwright() as pw:
            browser = await pw.chromium.launch(headless=True, args=['--disable-blink-features=AutomationControlled'])
            ctx = await browser.new_context(locale='en-US', viewport={'width': 1440, 'height': 900},
                                            user_agent=('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                                                        '(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'))
            if cookies:
                await ctx.add_cookies(cookies)
            s = Scraper(inp, ctx)
            await asyncio.gather(*(s.run_region(r, ms) for r, ms in by_region.items()))
            health.update(s.health)
            if inp.get('returnRefreshedSessions'):
                fresh = await ctx.cookies()
                root = {'US': 'amazon.com', 'EU': 'amazon.com', 'JP': 'amazon.co.jp', 'IN': 'amazon.in'}
                await Actor.set_value('REFRESHED_SESSIONS', {
                    r: [c for c in fresh if c['domain'].lstrip('.') in (REGIONS[r]['host'], root[r])]
                    for r in by_region})
                Actor.log.info('Refreshed sessions saved to the key-value store record REFRESHED_SESSIONS '
                               '(private to your account) — paste them into the session fields for the next run.')
            await browser.close()

        report = {'finishedAt': datetime.now(timezone.utc).isoformat(), 'reviews': s.pushed if cookies else 0,
                  'marketplaces': health}
        await Actor.set_value('SESSION_HEALTH', report)
        bad = {m: h['status'] for m, h in health.items() if h['status'] != 'OK'}
        Actor.log.info(f'Session health: {health}')
        if bad and len(bad) == len(markets):
            # User-side problem (expired/missing session) — report clearly, don't count as an actor failure.
            await Actor.exit(status_message=f'⚠️ No marketplace scraped: {bad}. Refresh your sessions and run again '
                                            f'(details: key-value store record SESSION_HEALTH).')
            return
        msg = f'{report["reviews"]} reviews' + (f' · needs attention: {bad}' if bad else ' · all sessions healthy')
        await Actor.exit(status_message=msg)

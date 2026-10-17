"""Saved Seller Central sessions, one per region, kept in a named key-value store in the RUNNING user's own
Apify account (private to them; another user of this public actor gets their own store and never sees these).

Record layout (key = SESSION_<REGION>):
    {"region": "US", "savedAt": "...Z", "source": "login-mode" | "pasted" | "refreshed", "cookies": [<playwright cookies>]}

Cookie values are never logged. Only names/counts are, via sessions.fingerprint().
"""
from datetime import datetime, timezone

from apify import Actor

DEFAULT_STORE = 'sc-review-sessions'

# Cookies that belong to a region's Seller Central login (host + the shared amazon root it signs in through).
REGION_DOMAINS = {
    'US': ('sellercentral.amazon.com', 'amazon.com'),
    'EU': ('sellercentral-europe.amazon.com', 'amazon.com', 'amazon.co.uk', 'amazon.de', 'amazon.fr', 'amazon.it',
           'amazon.es', 'sellercentral.amazon.de', 'sellercentral.amazon.co.uk', 'sellercentral.amazon.fr',
           'sellercentral.amazon.it', 'sellercentral.amazon.es'),
    'JP': ('sellercentral.amazon.co.jp', 'amazon.co.jp'),
    'IN': ('sellercentral.amazon.in', 'amazon.in'),
}


def store_name(inp: dict) -> str:
    name = (inp.get('sessionStoreName') or DEFAULT_STORE).strip().lower()
    return ''.join(ch if ch.isalnum() or ch == '-' else '-' for ch in name)[:63] or DEFAULT_STORE


def region_cookies(cookies: list[dict], region: str) -> list[dict]:
    doms = REGION_DOMAINS[region]
    out = []
    for c in cookies:
        d = (c.get('domain') or '').lstrip('.')
        if not d and c.get('url'):
            d = c['url'].split('/')[2]
        if any(d == x or d.endswith('.' + x) for x in doms):
            out.append(c)
    return out


def _clean(cookies: list[dict]) -> list[dict]:
    """Playwright's ctx.cookies() → the shape ctx.add_cookies() accepts again."""
    keep = ('name', 'value', 'domain', 'path', 'expires', 'httpOnly', 'secure', 'sameSite', 'url')
    out = []
    for c in cookies:
        x = {k: c[k] for k in keep if k in c and c[k] is not None}
        if x.get('expires') is not None and x['expires'] <= 0:
            x.pop('expires')     # session cookie
        if 'url' in x and 'domain' in x:
            x.pop('url')
        out.append(x)
    return out


async def load(inp: dict, region: str) -> dict | None:
    store = await Actor.open_key_value_store(name=store_name(inp))
    rec = await store.get_value(f'SESSION_{region}')
    return rec if isinstance(rec, dict) and rec.get('cookies') else None


async def save(inp: dict, region: str, cookies: list[dict], source: str) -> int:
    cookies = _clean(region_cookies(cookies, region))
    if not cookies:
        return 0
    store = await Actor.open_key_value_store(name=store_name(inp))
    await store.set_value(f'SESSION_{region}', {
        'region': region, 'source': source, 'cookies': cookies,
        'savedAt': datetime.now(timezone.utc).isoformat(timespec='seconds'),
    })
    return len(cookies)


async def forget(inp: dict, region: str):
    store = await Actor.open_key_value_store(name=store_name(inp))
    await store.set_value(f'SESSION_{region}', None)

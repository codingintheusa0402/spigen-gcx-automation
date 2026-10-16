"""Parse pasted Seller Central / Amazon login sessions into Playwright cookies.

Accepted formats (auto-detected):
  1. Cookie-Editor / EditThisCookie JSON export: [{"name", "value", "domain", ...}, ...]
  2. Playwright storage-state JSON: {"cookies": [...], "origins": [...]}
  3. Raw Cookie header copied from DevTools: "session-id=...; at-main=...; ..."
  4. Netscape cookies.txt lines (tab-separated)
"""
import json
import re

SAME_SITE = {'no_restriction': 'None', 'none': 'None', 'lax': 'Lax', 'strict': 'Strict', 'unspecified': 'Lax'}


def _from_json_cookie(c: dict) -> dict | None:
    name, value = c.get('name'), c.get('value')
    domain = c.get('domain') or c.get('host')
    if not name or value is None or not domain:
        return None
    out = {'name': str(name), 'value': str(value), 'path': c.get('path') or '/',
           'secure': bool(c.get('secure', True)), 'httpOnly': bool(c.get('httpOnly', False))}
    host_only = c.get('hostOnly')
    if name.startswith('__Host-') or host_only:
        out['url'] = f"https://{domain.lstrip('.')}{out['path']}"
        out.pop('path')
    else:
        out['domain'] = domain
    exp = c.get('expirationDate', c.get('expires'))
    if isinstance(exp, (int, float)) and exp > 0:
        out['expires'] = float(exp)
    ss = SAME_SITE.get(str(c.get('sameSite', '')).lower())
    if ss:
        out['sameSite'] = ss
        if ss == 'None':
            out['secure'] = True
    return out


def parse_session(text: str, default_host: str) -> list[dict]:
    """Return Playwright cookie dicts. `default_host` is used for raw header strings."""
    text = (text or '').strip()
    if not text:
        return []
    if text[0] in '[{':
        data = json.loads(text)
        if isinstance(data, dict):
            data = data.get('cookies', [])
        return [c for c in (_from_json_cookie(x) for x in data if isinstance(x, dict)) if c]
    if '\t' in text and re.search(r'^\S+\t(TRUE|FALSE)\t', text, re.M):  # Netscape cookies.txt
        out = []
        for line in text.splitlines():
            if line.startswith('#') or not line.strip():
                continue
            p = line.split('\t')
            if len(p) >= 7:
                c = _from_json_cookie({'domain': p[0], 'path': p[2], 'secure': p[3] == 'TRUE',
                                       'expirationDate': float(p[4] or 0), 'name': p[5], 'value': p[6]})
                if c:
                    out.append(c)
        return out
    # Raw "Cookie:" header — strip an optional leading "cookie:" label.
    text = re.sub(r'^\s*cookie\s*:\s*', '', text, flags=re.I)
    out = []
    for part in text.split(';'):
        if '=' not in part:
            continue
        name, value = part.split('=', 1)
        name, value = name.strip(), value.strip()
        if name:
            # Host-only on the Seller Central host: exactly where the browser sent them.
            out.append({'name': name, 'value': value, 'url': f'https://{default_host}/', 'secure': True})
    return out


def fingerprint(cookies: list[dict]) -> str:
    """Short, non-reversible description for logs (never log values)."""
    names = sorted({c['name'] for c in cookies})
    has_auth = any(n.startswith(('at-', 'sess-at-', 'x-', 'sst-')) or n in ('session-token',) for n in names)
    return f"{len(cookies)} cookies{'' if has_auth else ' (no auth cookies found — is this the right export?)'}"

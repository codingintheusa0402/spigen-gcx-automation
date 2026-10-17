#!/usr/bin/env python3
"""Substitute-product recommendation for a Zendesk ticket.

    python3 recommend.py https://spigenhelp.zendesk.com/agent/tickets/1000166420
    python3 recommend.py 1000166420 --sku ACS10066 --market DE      # override what the ticket says
    python3 recommend.py --sku ACS10066 --market GB --all           # no ticket; every category, not just the same kind

1. Ticket → product (★문의SKU / ASIN fields, else SKU/ASIN in subject+description) and the customer's
   marketplace (Point of Purchase / Country fields), read through the logged-in Zendesk Chrome (CDP :9222)
   — there is no Zendesk API token on these Macs.
2. Same-device, same-kind substitutes from the Chat app's own matcher (engine.js → ProductSearch.gs) over the
   ProductIndex tab; stock + listing state from the InventorySnapshot / ListingsSnapshot tabs (hourly / 6-hourly,
   EU countries + JP; NA is not covered by the MCF SP-API credentials).
3. Keeps substitutes with FBA stock in that country AND an active listing there, most stock first, and prints a
   copy-paste table (ASIN linked to that country's Amazon page | Device | Product | SKU). On macOS it also puts
   the table on the clipboard as rich text, so pasting into Zendesk gives a real table with links.
"""
import argparse, html, json, os, re, subprocess, sys, urllib.parse, urllib.request

HERE = os.path.dirname(os.path.realpath(__file__))
ENGINE = os.path.join(HERE, 'engine.js')
PRODUCT_SEARCH = os.path.realpath(os.path.join(HERE, '..', '..', '..', 'GAS_Operations', 'TicketReporterCard', 'ProductSearch.gs'))
QUEUE_SHEET = '1GUU2CLa60XSGWHI6yZ6Jkx4nrzcGgBJnuxwQAZktH7M'   # TicketQueue spreadsheet (snapshot tabs live here)
TOKEN_FILE = os.path.expanduser('~/.config/gws_shim/token.json')
CDP = 'http://localhost:9222'
ZD = 'https://spigenhelp.zendesk.com'
F_COUNTRY, F_POP, F_SKU, F_ASIN, F_DEVICE, F_PRODUCT = '4513936822297', '20016270875033', '900008676703', '360021934312', '360022185671', '360022185891'

DOMAINS = {'DE': 'amazon.de', 'GB': 'amazon.co.uk', 'FR': 'amazon.fr', 'IT': 'amazon.it', 'ES': 'amazon.es', 'NL': 'amazon.nl',
           'SE': 'amazon.se', 'PL': 'amazon.pl', 'BE': 'amazon.com.be', 'IE': 'amazon.ie', 'TR': 'amazon.com.tr', 'AE': 'amazon.ae',
           'SA': 'amazon.sa', 'JP': 'amazon.co.jp', 'US': 'amazon.com', 'CA': 'amazon.ca', 'MX': 'amazon.com.mx', 'IN': 'amazon.in',
           'AU': 'amazon.com.au', 'SG': 'amazon.sg', 'BR': 'amazon.com.br'}
COVERED = {'DE', 'GB', 'FR', 'IT', 'ES', 'NL', 'SE', 'PL', 'BE', 'IE', 'TR', 'AE', 'SA', 'JP'}   # snapshot countries
# words in Point of Purchase / Country values → country code (checked as whole words, longest first)
MARKET_WORDS = {
    'DE': ['amazon.de', 'germany', 'deutschland', 'de', '독일'], 'GB': ['amazon.co.uk', 'united kingdom', 'uk', 'gb', 'england', '영국'],
    'FR': ['amazon.fr', 'france', 'fr', '프랑스'], 'IT': ['amazon.it', 'italy', 'italia', 'it', '이탈리아'], 'ES': ['amazon.es', 'spain', 'españa', 'es', '스페인'],
    'NL': ['amazon.nl', 'netherlands', 'nl', '네덜란드'], 'SE': ['amazon.se', 'sweden', 'se', '스웨덴'], 'PL': ['amazon.pl', 'poland', 'pl', '폴란드'],
    'BE': ['amazon.com.be', 'belgium', 'be', '벨기에'], 'IE': ['amazon.ie', 'ireland', 'ie', '아일랜드'], 'TR': ['amazon.com.tr', 'turkey', 'türkiye', 'tr', '튀르키예', '터키'],
    'AE': ['amazon.ae', 'uae', 'united arab emirates', 'ae'], 'SA': ['amazon.sa', 'saudi', 'saudi arabia', 'sa'], 'JP': ['amazon.co.jp', 'japan', 'jp', '일본'],
    'US': ['amazon.com', 'united states', 'usa', 'us', '미국'], 'CA': ['amazon.ca', 'canada', 'ca', '캐나다'], 'MX': ['amazon.com.mx', 'mexico', 'mx', '멕시코'],
    'IN': ['amazon.in', 'india', 'in', '인도'], 'AU': ['amazon.com.au', 'australia', 'au', '호주'],
}
SKU_RE = re.compile(r'\b([A-Z]{3}\d{5}|\d{3}[A-Z]{2}\d{5})(?:[A-Z]{0,5})\b')
ASIN_RE = re.compile(r'\b(B0[A-Z0-9]{8})\b')


def die(msg):
    print('❌ ' + msg, file=sys.stderr)
    sys.exit(1)


# ---------------- Zendesk (through the logged-in agent Chrome) ----------------

def zendesk_get(paths):
    """GET several Zendesk API paths in the agent's Chrome session → list of parsed JSON (None on error)."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        try:
            b = p.chromium.connect_over_cdp(CDP)
        except Exception as e:
            die(f'Zendesk Chrome (CDP {CDP}) is not running: {e}')
        ctx = b.contexts[0]
        pg = next((x for x in ctx.pages if 'zendesk.com' in x.url), None)
        own = pg is None
        if own:
            pg = ctx.new_page()
            pg.goto(ZD + '/api/v2/users/me.json')
        res = pg.evaluate("""async (paths) => Promise.all(paths.map(async u => {
            const r = await fetch(u, {credentials: 'include'}); return [r.status, r.ok ? await r.json() : null]; }))""",
                          [ZD + x for x in paths])
        if own:
            pg.close()
    if res and res[0][0] == 401:
        die('Zendesk Chrome is not signed in — sign in at spigenhelp.zendesk.com in the :9222 Chrome (Screen Sharing), then retry.')
    return [j for _, j in res]


def field_name(fields_meta, fid, value):
    """Tagger value → its display name (falls back to the raw value)."""
    for opt in (fields_meta.get(fid) or {}).get('custom_field_options', []):
        if opt.get('value') == value:
            return opt.get('name') or value
    return value


def read_ticket(tid):
    tk, *metas = zendesk_get([f'/api/v2/tickets/{tid}.json'] + [f'/api/v2/ticket_fields/{f}.json' for f in (F_COUNTRY, F_POP, F_SKU)])
    if not tk:
        die(f'ticket {tid} could not be read')
    tk = tk['ticket']
    meta = {f: (m or {}).get('ticket_field', {}) for f, m in zip((F_COUNTRY, F_POP, F_SKU), metas)}
    cf = {str(f['id']): f['value'] for f in tk.get('custom_fields', []) if f.get('value') not in (None, '', [])}
    vals = {k: field_name(meta, k, v) if isinstance(v, str) else v for k, v in cf.items()}
    return tk, vals


def market_from(*texts):
    for t in texts:
        t = ' ' + re.sub(r'[^\w.\s가-힣]', ' ', str(t or '').lower()) + ' '
        hits = [(len(w), c) for c, ws in MARKET_WORDS.items() for w in ws if (' ' + w + ' ') in t or ('.' in w and w in t)]
        if hits:
            return max(hits)[1]
    return None


# ---------------- Snapshots (Sheets API, gws_shim token) ----------------

def google_token():
    t = json.load(open(TOKEN_FILE))
    d = urllib.parse.urlencode({'client_id': t['client_id'], 'client_secret': t['client_secret'],
                                'refresh_token': t['refresh_token'], 'grant_type': 'refresh_token'}).encode()
    return json.load(urllib.request.urlopen('https://oauth2.googleapis.com/token', d))['access_token']


def read_tab_json(tok, tab):
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{QUEUE_SHEET}/values/{urllib.parse.quote(tab)}!A:A'
    rows = json.load(urllib.request.urlopen(urllib.request.Request(url, headers={'Authorization': 'Bearer ' + tok}))).get('values', [])
    return json.loads(''.join(r[0] for r in rows if r))


def merged(p, snap):
    """Stock / listing entry for a product: its SKU key plus "ASIN:" keys (legacy seller SKUs)."""
    out = {'c': {}, 'on': set(), 'ac': {}}
    for k in [p['sku']] + ['ASIN:' + a for a in p.get('asins', [])]:
        o = (snap or {}).get('by', {}).get(k)
        if not o:
            continue
        for c, q in (o.get('c') or {}).items():
            out['c'][c] = out['c'].get(c, 0) + q
        out['on'] |= set((o.get('on') or '').split())
        for c, a in (o.get('ac') or {}).items():
            out['ac'].setdefault(c, a)
    return out


# ---------------- Output ----------------

def to_clipboard(html_table, plain):
    if sys.platform != 'darwin':
        return False
    hexd = html_table.encode('utf-8').hex()
    script = f'set the clipboard to {{«class HTML»:«data HTML{hexd}», string:{json.dumps(plain)}}}'
    try:
        subprocess.run(['osascript', '-e', script], check=True, capture_output=True, timeout=15)
        return True
    except Exception:
        return False


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('ticket', nargs='?', help='Zendesk ticket URL or number')
    ap.add_argument('--sku', help='product to replace (overrides the ticket)')
    ap.add_argument('--asin', help='product to replace, by ASIN')
    ap.add_argument('--market', help='customer country code, e.g. DE, GB/UK, JP (overrides the ticket)')
    ap.add_argument('--all', action='store_true', help='every category for the device, not just the same kind')
    ap.add_argument('--n', type=int, default=10, help='rows (default 10)')
    ap.add_argument('--no-copy', action='store_true', help="don't put the table on the macOS clipboard")
    ap.add_argument('--json', action='store_true')
    a = ap.parse_args()

    ctx = {}
    if a.ticket:
        tid = re.search(r'(\d{6,})', a.ticket)
        if not tid:
            die('no ticket number in ' + a.ticket)
        tk, vals = read_ticket(tid.group(1))
        text = ' '.join([tk.get('subject') or '', tk.get('description') or '']).upper()
        ctx = {'ticket': tid.group(1), 'subject': tk.get('subject'), 'fields': {k: vals.get(k) for k in (F_COUNTRY, F_POP, F_SKU, F_ASIN, F_DEVICE, F_PRODUCT)}}
        sku_field = str(vals.get(F_SKU) or '').upper()
        a.sku = a.sku or (SKU_RE.search(sku_field) or SKU_RE.search(text) or [None, None])[1]
        a.asin = a.asin or (ASIN_RE.search(str(vals.get(F_ASIN) or '').upper()) or ASIN_RE.search(text) or [None, None])[1]
        a.market = a.market or market_from(vals.get(F_POP), vals.get(F_COUNTRY))
    if not (a.sku or a.asin):
        die('no SKU / ASIN found' + (' on the ticket (★문의SKU / ASIN fields empty) — pass --sku' if a.ticket else ' — pass --sku or --asin'))
    if not a.market:
        die('customer marketplace unknown (Point of Purchase / Country empty) — pass --market DE|GB|FR|IT|ES|JP|…')
    mk = a.market.upper().replace('UK', 'GB')
    if mk not in DOMAINS:
        die('unknown marketplace ' + a.market)

    tok = google_token()
    index = read_tab_json(tok, 'ProductIndex')
    inv = read_tab_json(tok, 'InventorySnapshot')
    lst = read_tab_json(tok, 'ListingsSnapshot')
    by_sku = {p['sku']: p for p in index}

    query = (a.sku or a.asin) + (' all' if a.all else '')
    r = subprocess.run(['node', ENGINE, PRODUCT_SEARCH], input=json.dumps({'index': index, 'query': query}),
                       capture_output=True, text=True, check=True)
    sim = json.loads(r.stdout)
    if sim.get('error'):
        die(re.sub(r'^⚠️\s*', '', sim['error']))
    base = by_sku.get(sim['base'], {})

    rows = []
    for sku in sim['same']:
        p = by_sku[sku]
        s, l = merged(p, inv), merged(p, lst)
        if mk in COVERED:
            stock = s['c'].get(mk, 0)
            asin = l['ac'].get(mk) or s['ac'].get(mk)
            if stock <= 0 or mk not in l['on'] or not asin:
                continue
        else:   # NA etc.: no stock / listing data — candidates only, master ASIN
            stock, asin = None, (p.get('asins') or [None])[0]
            if not asin:
                continue
        rows.append({'asin': asin, 'device': re.sub(r'^#|/+$', '', p.get('device', '')), 'sku': sku, 'stock': stock,
                     'product': ' '.join(x for x in (p.get('brand'), p.get('model'), p.get('color')) if x)})
    rows.sort(key=lambda x: -(x['stock'] or 0))
    rows = rows[:a.n]

    if a.json:
        print(json.dumps({'context': ctx, 'market': mk, 'base': sim['base'], 'device': sim['label'], 'kind': sim['cats'], 'rows': rows}, ensure_ascii=False, indent=1))
        return

    dom = DOMAINS[mk]
    link = lambda x: f'https://www.{dom}/dp/{x}'
    head = (f"#{ctx['ticket']} · " if ctx else '') + f"{mk} ({dom}) · 기준 {sim['base']} {base.get('model', '')} {base.get('color', '')}".rstrip()
    kind = ' · '.join(sim['cats']) if sim['cats'] else '전체 카테고리'
    print(f'{head}\n기종 {sim["label"]} · {kind} · 같은 기종 후보 {len(sim["same"])}건 → '
          + (f'{mk} 재고·활성 리스팅 있는 {len(rows)}건' if mk in COVERED else f'{mk} 재고 확인 불가 (NA는 SP-API 미지원) — 후보만'))
    if not rows:
        print('\n(추천할 대체 상품 없음 — 같은 기종·같은 종류 중 이 마켓에 재고+활성 리스팅이 있는 제품이 없습니다. --all 로 전체 카테고리를 볼 수 있습니다)')
        return
    md = ['| ASIN | Device | Product | SKU |', '|---|---|---|---|']
    md += [f"| [{x['asin']}]({link(x['asin'])}) | {x['device']} | {x['product']} | {x['sku']} |" for x in rows]
    print('\n' + '\n'.join(md))
    if mk in COVERED:
        print(f"\n(재고 순 · {mk} FBA 판매가능: " + ', '.join(f"{x['sku']} {x['stock']:,}" for x in rows) + ')')

    if not a.no_copy:
        th = ''.join(f'<th style="border:1px solid #ccc;padding:4px 8px;text-align:left">{h}</th>' for h in ('ASIN', 'Device', 'Product', 'SKU'))
        trs = ''.join('<tr>' + ''.join(f'<td style="border:1px solid #ccc;padding:4px 8px">{c}</td>' for c in (
            f'<a href="{link(x["asin"])}">{x["asin"]}</a>', html.escape(x['device']), html.escape(x['product']), x['sku'])) + '</tr>' for x in rows)
        table = f'<table style="border-collapse:collapse"><tr>{th}</tr>{trs}</table>'
        plain = '\n'.join(['ASIN\tDevice\tProduct\tSKU'] + [f"{x['asin']} ({link(x['asin'])})\t{x['device']}\t{x['product']}\t{x['sku']}" for x in rows])
        if to_clipboard(table, plain):
            print('\n📋 클립보드에 표(링크 포함)가 복사되었습니다 — Zendesk 답변창에 바로 붙여넣기 하세요.')


if __name__ == '__main__':
    main()

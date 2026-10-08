"""Step 7 — post one review-request card per case to the PRIVATE room (format of the 2026-09-30 Pixel card +
'SIREN 덱 열기' button). Long cases continue in the same thread. Usage: python3 07_send_cards.py 1 2 3 ...
Never send to a team room unless the user says so."""
"""Posts each case's SIREN review-request card (same format as the 2026-09-30 Pixel card) to the
private room, with the deck link. Long cases continue in the same thread."""
import sys, json, re, urllib.request, time
import common  # noqa
WEBHOOK = common.PRIVATE_WEBHOOK

def post(payload, thread_key):
    url = WEBHOOK + f"&threadKey={thread_key}&messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD"
    req = urllib.request.Request(url, data=json.dumps(payload, ensure_ascii=False).encode(),
                                 headers={"Content-Type": "application/json; charset=UTF-8"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read())["name"]

REG = json.load(open('regmatch.json')) if __import__('os').path.exists('regmatch.json') else {}
def note(n):
    """Red line on the card when the case is already a registered SIREN (from 06_registry_check)."""
    m = REG.get(str(n))
    if not m: return None
    sure = [r for r in m if not r['possible']]
    if sure:
        r = sure[0]
        return f"※ 기등록 SIREN: {r['src']} ({r['date']}, {r['issue'][:40]}) — 등록 후에도 VOC 지속"
    r = m[0]
    return f"※ 기등록 가능성: {r['src']} ({r['date']}, {r['issue'][:40]}) — SKU 일부만 기재되어 확인 필요"

def send(n):
    d = json.load(open(f'case_{n:02d}.final.json'))
    deck = open(f'deck_{n:02d}.txt').read().strip()
    p = d['product']; name = p['product_name']; defect = p['defect_type']
    claims = [c for c in d['claims'] if c['label'] == '클레임']
    reviews = [c for c in d['claims'] if c['label'] != '클레임']
    head = (f"리더님, 프로님들 안녕하세요. {name} {defect} 이슈 클레임 {len(claims)}건 "
            f"배드리뷰 {len(reviews)}건이 발견되어 SIREN등록 해도 되는지 검토 부탁드립니다. 감사합니다!")
    lines = []
    for i, c in enumerate(claims, 1):
        a = c.get('_asin') or p['asin']
        lines.append(f'<b><a href="https://spigenhelp.zendesk.com/agent/tickets/{c["link_label"]}">[클레임{i}]</a></b>')
        lines.append(f'{name} [<a href="https://www.amazon.de/dp/{a}">{a}</a>] {defect} 이슈 ({c["country"]})')
    if reviews:
        lines.append('')
        for a in sorted({r.get('_asin') or p['asin'] for r in reviews}):
            lines.append(f'{name} [<a href="https://www.amazon.de/dp/{a}">{a}</a>] {defect} 이슈')
            for i, r in enumerate(reviews, 1):
                if (r.get('_asin') or p['asin']) == a:
                    url = r.get('link_url') or ''
                    lines.append(f'<b><a href="{url}">[배드리뷰{i}]</a></b> ({r["country"]})')
    images = [u for c in d['claims'] for u in c.get('image_urls', [])]
    # split the list so each message stays well under Chat's 32KB limit
    parts, cur = [], []
    for ln in lines:
        if len('<br>'.join(cur + [ln]).encode()) > 18000:
            parts.append(cur); cur = []
        cur.append(ln)
    parts.append(cur)
    key = f"siren{__import__('datetime').date.today():%y%m%d}case{n:02d}"
    names = []
    for k, part in enumerate(parts):
        widgets = []
        if k == 0:
            nt = note(n)
            note_html = f'<font color="#d93025"><b>{nt}</b></font><br><br>' if nt else ''
            widgets.append({"textParagraph": {"text": f"<b>[SIREN 후보 {n}/{TOTAL}]</b><br>" + note_html + head + "<br><br>" + "<br>".join(part)}})
        else:
            widgets.append({"textParagraph": {"text": f"(계속 {k + 1}/{len(parts)})<br>" + "<br>".join(part)}})
        if k == len(parts) - 1:
            if images:
                capped = images[:10]
                widgets.append({"textParagraph": {"text": f"고객 첨부 사진 ({len(images)}장)"}})
                widgets.append({"grid": {"columnCount": 3 if len(capped) >= 3 else len(capped),
                                         "items": [{"id": str(i), "image": {"imageUri": u}} for i, u in enumerate(capped)]}})
                if len(images) > 10:
                    widgets.append({"textParagraph": {"text": f"…외 {len(images) - 10}건 이미지 생략 (전체는 SIREN 덱 참고)"}})
            widgets.append({"buttonList": {"buttons": [{"text": "SIREN 덱 열기", "onClick": {"openLink": {"url": deck}}}]}})
        payload = {"cardsV2": [{"cardId": f"siren-review-request-{n:02d}-{k}", "card": {"sections": [{"widgets": widgets}]}}]}
        names.append(post(payload, key))
        time.sleep(1.2)
    return names

TOTAL = len(__import__('glob').glob('case_[0-9][0-9].json'))

if __name__ == '__main__':
    for n in map(int, sys.argv[1:]):
        print(n, send(n), flush=True)

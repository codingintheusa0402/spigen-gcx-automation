"""Step 13 — post the summary deck card (취지/범위 + '자료 열기') to the PRIVATE room only.
Manager rule: always state the purpose (취지/목적) when sharing. Usage: python3 13_send_summary_card.py [deck_id]"""
import sys, json, datetime, urllib.request, common  # noqa
pid = sys.argv[1] if len(sys.argv) > 1 else open('summary_deck.txt').read().strip()
cases = json.load(open('cases.json')); regs = [c for c in cases if c['reg']]
half = common.half_label(); title = f'GCX SIREN {half} 등록현황 및 VOC 점검'
text = (f"「{title}」 공유드립니다.<br><br>"
        "<b>취지:</b> SIREN 등록 누락을 막기 위해, 동일 라인업 × 동일 불량으로 클레임+배드리뷰가 5건 이상 쌓인 이슈를 정기적으로 점검하고, "
        "이미 등록된 이슈도 등록 이후 VOC가 계속 들어오는지 함께 확인하는 자료입니다.<br>"
        f"<b>범위:</b> Screen Protector(글라스) 제외 {len(cases)}건 (글라스는 별도 팀 보고)."
        + (f" 이 중 {len(regs)}건은 기등록이지만 VOC가 계속 인입되고 있습니다." if regs else ''))
payload = {"cardsV2": [{"cardId": "siren-sweep-summary", "card": {
    "header": {"title": title, "subtitle": f"{datetime.date.today():%Y.%m.%d} · 글로벌CX전략팀"},
    "sections": [{"widgets": [{"textParagraph": {"text": text}},
        {"buttonList": {"buttons": [{"text": "자료 열기", "onClick": {"openLink": {"url": f"https://docs.google.com/presentation/d/{pid}/edit"}}}]}}]}]}}]}
req = urllib.request.Request(common.PRIVATE_WEBHOOK, data=json.dumps(payload, ensure_ascii=False).encode(), headers={"Content-Type": "application/json; charset=UTF-8"})
print(json.loads(urllib.request.urlopen(req, timeout=60).read())['name'])

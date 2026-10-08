"""Step 8 — post the 'registered SIREN but VOC keeps coming' report (from regreport.json) to the PRIVATE room.
Criteria: VOC after the latest registration >= 10, or >= 5 in the last 30 days. Usage: python3 08_send_reg_report.py"""
import json, re, time, datetime, urllib.request, common  # noqa
items = json.load(open('regreport.json'))
hot = [i for i in items if i['after'] >= 10 or i['last30'] >= 5]
S26 = f"https://docs.google.com/spreadsheets/d/{common.GCX_KPI_26[0]}/edit"
S25 = f"https://docs.google.com/spreadsheets/d/{common.GCX_KPI_25}/edit"
SAL = f"https://docs.google.com/spreadsheets/d/{common.SALES_SIREN[0]}/edit"
cases_reg = json.load(open('regmatch.json')) if __import__('os').path.exists('regmatch.json') else {}
imp_still = sum(1 for i in hot if any(r['imp'].strip().upper().startswith('O') or '완료' in r['imp'] for r in i['regs']))
head = (f"*[기등록 SIREN — VOC 지속 리포트]* {datetime.date.today():%Y-%m-%d}\n"
        f"대상: <{S26}|GCX SIREN 등록 현황> + <{S25}|전년 KPI 실적 검증> + <{SAL}|SIREN 등록 현황(영업)>\n"
        f"방법: 같은 SKU 라인업 × 같은 불량(Zendesk 클레임 + 아마존 1~3점 리뷰, 올해)을 *가장 최근 SIREN 등록일 이후* 로 집계\n"
        f"기준: 등록 후 VOC 10건 이상 또는 최근 30일 5건 이상 → *{len(hot)}건* (그중 개선 완료/O 인데도 지속 *{imp_still}건*)\n"
        + (f"※ 이번 SIREN 후보 중 기등록: {', '.join('#' + n for n in sorted(cases_reg, key=int))}\n" if cases_reg else ''))
lines = []
for n, i in enumerate(hot, 1):
    r = i['regs'][-1]
    regs = ", ".join(f"{x['src']}({x['date']}, 개선 {x['imp'][:10] or '-'})" for x in i['regs'])
    res = re.sub(r'\s+', ' ', r['result'])[:70]
    flag = " 🔴개선 후에도 지속" if any(x['imp'].strip().upper().startswith('O') or '완료' in x['imp'] for x in i['regs']) else ""
    links = " ".join(f"<{u}|{'#' if 'zendesk' in u else 'R'}{k + 1}>" for k, u in enumerate(i['links']))
    lines.append(f"*{n}. {r['product'][:45]}* ({', '.join(r['skus'][:3]) or '-'}) — {r['issue'][:25]}{flag}\n"
                 f"    등록: {regs}\n"
                 f"    등록 후 VOC *{i['after']}건* (클레임 {i['claims']} · 리뷰 {i['reviews']}) · 최근 30일 *{i['last30']}건* · 최근 인입 {i['last']}\n"
                 + (f"    개선 결과: {res}\n" if res else "") + f"    최근 사례: {links}")
msgs, cur = [], head
for ln in lines:
    if len(cur) + len(ln) + 2 > 3800: msgs.append(cur); cur = "(계속)"
    cur += "\n" + ln
msgs.append(cur)
key = f"sirenregvoc{datetime.date.today():%y%m%d}"
for m in msgs:
    req = urllib.request.Request(common.PRIVATE_WEBHOOK + f"&threadKey={key}&messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD",
                                 data=json.dumps({"text": m}, ensure_ascii=False).encode(), headers={"Content-Type": "application/json; charset=UTF-8"})
    print(json.loads(urllib.request.urlopen(req, timeout=60).read())['name']); time.sleep(1.2)

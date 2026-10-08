"""Step 3 — Korean 상세 내용 (1–2 sentences) for every claim ticket via `claude -p --model sonnet`, 25/batch, 8 parallel.
Resumable (ko.json cache). Re-run until it reports all tickets summarised."""
import sys, json, glob, os, concurrent.futures as cf
import common  # noqa  (chdir to work dir)
import find_siren as F
recs = json.load(open('tickets.json'))
OUT = 'ko.json'
cache = json.load(open(OUT)) if os.path.exists(OUT) else {}
ctx = {}
for f in sorted(glob.glob('case_*.json')):
    d = json.load(open(f))
    for c in d['claims']:
        if c['label'] == '클레임':
            ctx[c['link_label']] = (d['product']['product_name'], d['product']['defect_type'], c['detail_text'])
todo = [i for i in ctx if i not in cache]
PROMPT = """아래는 Spigen 고객 클레임(Zendesk 티켓) 원문 목록이다. 각 티켓마다 SIREN 보고서 '상세 내용' 칸에 들어갈
한국어 요약을 1~2문장(최대 150자)으로 작성하라. 문체 예시:
"Google Pixel 10용 EZ Fit Privacy 글라스 부착 후 화면 전체에 분홍색/연한 파란색 줄무늬 현상 발생. 해결 방법 문의."
"Galaxy Z Fold7용 Tough Armor Pro MagFit 케이스 힌지 부분 지지대 4개 중 1개가 부러져 힌지 커버가 일부 탈락함."
- 고객이 말한 제품·증상·경과(언제/어떻게)·요청(교환/환불 등)만 사실대로. 추측·인사말 금지. 명사형 종결("~발생.", "~함.").
- 원문에 제품명이 없으면 주어진 product 를 쓴다.
JSON 배열만 반환: [{"id":"<id>","ko":"<요약>"}]

%s"""
batches = [todo[i:i+25] for i in range(0, len(todo), 25)]
def run(b):
    payload = [{"id": i, "product": ctx[i][0], "defect": ctx[i][1], "sku": ctx[i][2].strip(), "text": recs[i]['b'][:700]} for i in b]
    return F.claude_json(PROMPT % json.dumps(payload, ensure_ascii=False))
with cf.ThreadPoolExecutor(8) as ex:
    futs = [ex.submit(run, b) for b in batches]
    for n, f in enumerate(cf.as_completed(futs), 1):
        try:
            for it in f.result():
                if it.get('id') in ctx and it.get('ko'):
                    cache[it['id']] = it['ko']
        except Exception as e:
            print('batch failed', e, flush=True)
        json.dump(cache, open(OUT, 'w'), ensure_ascii=False)
        print(f'{n}/{len(batches)} done, {len(cache)} summaries', flush=True)

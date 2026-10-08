---
name: mcf-fee-backfill
description: "MCF 발송 로그 시트의 Transportation Fee (Y열) 중 MCFFee()/backfillMCFFeesRecent()/backfillMCFFees() 등 기존 자동화가 시도했지만 채우지 못한 값들을, 각 마켓플레이스(EU, JP, PowerArc 등) Seller Central의 /payments 정산 리포트를 수동으로 요청·다운로드해서 채운다. 사용자가 'MCF발송로그 Transaction Fee (Y col)', 'MCF Y컬럼 채워줘', 'MCF 정산 리포트로 채워줘' 등으로 요청할 때 발동."
license: MIT
metadata:
  category: automation
  locale: ko-KR
  phase: v1.0.0
---

# mcf-fee-backfill

MCF 발송 로그 시트 (`https://docs.google.com/spreadsheets/d/1g6a-S7eeA1oY19aTEFhTNAyp2A5nLqLNkPRqOqriWfc/edit?gid=1608794212`, "MCF 발송 로그" 탭)의
**Transportation Fee (Y열)** 이 비어있는 행 중, 기존 자동화(`MCFFee()`, `backfillMCFFeesRecent()`, `backfillMCFFees()` — `GAS_Operations/MCF_Tracking/` 프로젝트, scriptId
`1kDfEUVEEJ7TCA3HOMF6EYFTjbeZZKIeg_X84wCbLT1-tQqJI2ZlPUCxp`)가 이미 시도했지만 못 채운 것들을, 각 마켓플레이스 Seller Central의
Payments → All Statements 화면에서 정산 리포트를 **수동으로 요청·다운로드**해서 채우는 스킬. 이 세션에서 EU/JP를 실제로 채우며 얻은 시행착오가
전부 녹아있으니, 아래 순서를 그대로 따르면 다시 삽질하지 않는다.

## 0. 시작 전 반드시 확인할 것

- **gviz CSV export로 시트를 스캔하지 말 것** — 이유 없이 조용히 잘린다(한 번은 3758행 중 1750행만 반환). 대신 `MCF_Tracking` GAS 프로젝트에
  `_diag.js` 같은 1회성 진단 함수를 넣고 `SpreadsheetApp`으로 직접 읽는다 (아래 6번 참고).
- 시트 상수: `BF_SHEET_NAME`, `BF_START_ROW=4`, `BF_COL_REGION=2(B)`, `BF_COL_ORDER=17(Q, GCX- 주문ID)`,
  `BF_COL_SENT=16(P)`, `BF_COL_FEE=25(Y)`, 트래킹번호=col 18(R).
- **채우기 시도 전 이미 알려진 영구 불가 케이스**를 먼저 걸러낸다 (아래 5번). 여기 재조사에 시간 쓰지 말 것.

## 1. 대상 행 스코핑

1. `MCF_Tracking` GAS 프로젝트에 진단 함수를 추가해 `region`, `sentDate`, `orderId`, 현재 Y값을 스캔한다. Y가 빈칸이거나 `'RETRY'`이거나
   에러값(`_isErrorValue()`)인 행만 대상.
2. 대상 행을 **90일 이내 / 90일 초과**로 나눈다.
   - 90일 이내: 이미 `backfillMCFFeesRecent()`가 30분마다 자동으로 SP-API 정산 리포트 리스팅(`createdSince` ≤ ~89일)을 스캔 중.
     이 경우 먼저 `_loadFeeCache()` (persistent fee cache 시트, `_SettlementFeeCache`)에 해당 주문ID가 있는지 확인 —
     없으면 "아직 Amazon이 정산 안 함"이 원인일 확률이 높음 (버그 아님, 자동으로 채워질 것). 있는데 안 써졌으면 진짜 버그이니 별도 조사.
   - 90일 초과: SP-API 리스팅 엔드포인트의 `createdSince` 90일 제한 때문에 자동화가 원천적으로 못 건드림 → 이 스킬의 수동 프로세스 필요.

## 2. 마켓플레이스별 자격 증명

| 마켓플레이스 | credential 키 (`sp_get(..., cred=)`) | region | endpoint | marketplaceId |
|---|---|---|---|---|
| EU (DE/FR/IT/ES/**UK 포함**) | `main` | `eu-west-1` | `sellingpartnerapi-eu.amazon.com` | `A1PA6795UKMFR9` (DE 대표값, EU 전체 통합됨) |
| JP | `jp` | `us-west-2` | `sellingpartnerapi-fe.amazon.com` | `A1VC38T7YXB528` |
| PowerArc | **아직 없음** — 아래 "PowerArc 특이사항" 참고 | | | |

자격 증명 파일: `~/.sp-api-config.json`. 재사용 가능한 서명 헬퍼:
`~/Desktop/GCX/GAS_Zendesk/GCXReply_GAS/sp-api-proxy.py`의 `sp_get(endpoint, region, path, params, cred)` —
아래처럼 import (실제 FastAPI 서버가 뜨는 걸 막기 위해 `uvicorn.run`을 mock):

```python
import importlib.util, unittest.mock as mock
spec = importlib.util.spec_from_file_location('sp_api_proxy', '/Users/kevinkim/Desktop/GCX/GAS_Zendesk/GCXReply_GAS/sp-api-proxy.py')
mod = importlib.util.module_from_spec(spec)
with mock.patch('uvicorn.run'):
    spec.loader.exec_module(mod)
r = mod.sp_get('https://sellingpartnerapi-eu.amazon.com', 'eu-west-1', '/reports/2021-06-30/reports', params, 'main')
```

## 3. Seller Central에서 정산 리포트 요청 (브라우저 자동화)

각 마켓플레이스마다 반복:

1. 로그인된 Chrome 창을 찾는다 — `osascript`로 여러 창을 순회하며 `sellercentral`/`script.google.com` 탭이 있는 창을 찾을 것
   (탭 id/윈도우 id는 매번 바뀌므로 하드코딩 금지, 매번 새로 검색). **"Executing JavaScript through Apple Events가 꺼져있다"는
   에러가 나면 다른(잘못된) Chrome 창을 잡은 것** — 세션 내내 JS 실행이 됐던 창을 다시 찾아라.
2. 필요시 계정 전환: `https://sellercentral.amazon.de/account-switcher/default/merchantMarketplace` 에서 계정명(PowerArc/Spigen Direct/
   Spigen EU/Spigen Inc) 클릭 → 하위 국가 클릭 → `kat-button[label="Select account"]`의 shadow root 내부 button 클릭.
3. `https://sellercentral.{amazon.co.uk|amazon.co.jp|...}/payments/past-settlements` 로 이동.
4. "Welcome to your new All Statements page!" 투어 팝업이 뜨면 텍스트 `'Skip'` 매칭 요소를 풀 포인터 이벤트 시퀀스로 클릭해서 닫는다.
5. **From/To 날짜 필드 설정 — 여기가 제일 까다로운 부분, 아래 절차를 정확히 따를 것:**
   - `kat-date-picker`는 **일반 JS `.click()` + `.value=` 로 절대 안 먹는다** (내부 상태가 안 바뀜, Search 눌러도 무시됨).
     반드시 진짜 OS 레벨 클릭+타이핑이 필요하다.
   - `cliclick` 설치 확인: `which cliclick || brew install cliclick`.
   - **클릭 전에 반드시 해당 Chrome 창을 최전면으로 가져오고 올바른 탭을 활성화**:
     ```applescript
     tell application "Google Chrome"
       activate
       set win to window id <WINID>
       set index of win to 1
       set idx to 0
       set i to 0
       repeat with t in tabs of win
         set i to i + 1
         if (URL of t) contains "past-settlements" then set idx to i
       end repeat
       set active tab index of win to idx
     end tell
     ```
     이 단계를 생략하면 클릭이 엉뚱한 창으로 감 (이번 세션에서 여러 번 발생).
   - 좌표는 **스크린샷에서 눈대중으로 찍지 말고** JS로 정확히 계산:
     ```js
     var dps = document.querySelectorAll('kat-date-picker');
     var from = dps[0].shadowRoot.querySelector('kat-input').shadowRoot.querySelector('input');
     var to   = dps[1].shadowRoot.querySelector('kat-input').shadowRoot.querySelector('input');
     var fr = from.getBoundingClientRect(), tr = to.getBoundingClientRect();
     // Search 버튼: kat-button 중 getAttribute('label')이 /search/i 매칭하는 것
     JSON.stringify({from:{x:fr.x,y:fr.y}, to:{x:tr.x,y:tr.y}, screenX:window.screenX, screenY:window.screenY,
                      outerH:window.outerHeight, innerH:window.innerHeight});
     ```
     실제 화면 좌표 = `screenX + rect.x + rect.width/2` (x), `screenY + (outerH-innerH) + rect.y + rect.height/2` (y).
     `outerH-innerH`는 브라우저 툴바/탭바 높이 (보통 ~87pt). 멀티모니터/Retina 환경에서 스크린샷 픽셀 좌표를 그대로 쓰면 다른 모니터나
     엉뚱한 좌표를 클릭하게 되니 반드시 이 계산식을 쓸 것.
   - 클릭 후 **매번 포커스가 제대로 갔는지 확인** (`document.activeElement`가 어느 `kat-date-picker` index인지), 타이핑 후
     **입력값이 실제로 박혔는지 확인** (`input.value` 읽기) — 확인 없이 다음 동작으로 넘어가면 엉뚱한 필드에 타이핑되는 경우가 있었음.
   - `cliclick c:X,Y` → 포커스 확인 → `cliclick t:MM/DD/YYYY` → 값 확인, From/To 둘 다 반복 → Search 버튼도 같은 좌표 계산식으로 클릭.
6. 필터링된 리스트는 **최근 10개 기간만** 보여준다. 필요한 기간이 더 과거면 "To" 날짜를 더 이전으로 좁혀서 재검색 반복.
7. 각 기간 행의 Download 셀 확인:
   - `kat-dropdown-button` 있으면 → 이미 생성됨, `data-action` 속성이 바로 SP-API `reportId`. 그 안의
     `button.button` (텍스트 "Download Flat File V2")을 눌러도 되지만, **브라우저 다운로드는 신뢰하지 말고** (합성 클릭으로는
     실제 파일 다운로드가 안 트리거되는 경우가 잦았음) 그냥 `data-action` 값(reportId)만 뽑아서 아래 4번의 API 경로로 바로 fetch.
   - `kat-button` (라벨이 EU는 `"Request report"`, JP는 `"Request Report"` — **대소문자 다름**, 정규식 `/request report/i` 로 매칭)
     있으면 → 아직 생성 안 됨, shadow root 안의 `<button>`을 풀 포인터 이벤트 시퀀스(pointerdown/mousedown/pointerup/mouseup/click)로
     클릭해서 생성 요청. 필요한 행 전부 한 번에 클릭해도 됨 (배치 가능).

## 4. 리포트 완료 대기 + 다운로드 (API, 브라우저 아님)

- 방금 요청한 리포트들은 `createdTime`이 "지금"이 되므로, SP-API 리스팅을
  `createdSince=<2~3시간 전>` 으로 쿼리하면 90일 제한과 무관하게 바로 찾을 수 있다 (리스팅의 90일 제한은 **과거 커버리지** 필터링에 걸리는 것이지
  **생성 시각** 필터링엔 안 걸림).
- 백그라운드 Bash task로 30초 간격 폴링 (ScheduleWakeup으로 팔로우업 예약, 절대 `sleep 90` 같은 긴 슬립 직접 쓰지 말 것 — 차단됨).
- **429 QuotaExceeded가 자주 뜬다** — `reports: []`로 빈 리스트가 오는 게 아니라 `{"errors":[...]}` 형태라 `.get('reports',[])`가
  조용히 0을 리턴함. "리포트 개수가 갑자기 0으로 줄었다"고 당황하지 말고 몇 초 후 재시도하면 정상 데이터 옴.
- 일부 요청 건은 10분 넘게 기다려도 리포트가 안 생기는 경우가 있었음(이번 세션 JP 16개 중 2개). 무한 대기하지 말고,
  All Statements 페이지에서 해당 행이 아직도 `kat-button`("Request Report")이면 재요청, `kat-dropdown-button`으로 바뀌었으면
  API로 다시 조회. 그래도 끝까지 안 되면 그 주문은 "수동 조사 필요"로 남기고 넘어간다 (전체 작업을 막지 말 것).
- 리포트 다운로드:
  ```
  GET /reports/2021-06-30/reports/{reportId}  → reportDocumentId
  GET /reports/2021-06-30/documents/{reportDocumentId}  → { url, compressionAlgorithm }
  requests.get(url) → gzip.decompress if compressionAlgorithm=='GZIP'
  ```
  `/documents/{id}` 엔드포인트도 버스트 상황에서 429 잦음 — 재시도 backoff 걸 것 (10s, 20s, 30s...), 요청 사이 ~3초 pacing.
- 파싱: 탭 구분, 헤더를 **컬럼명으로** 인덱싱 (`merchant-order-id`, `amount-type`, `amount-description`, `amount` — EU/JP 공통 스키마).
  `merchant-order-id`가 `GCX-`로 시작하고 `amount-type=='ItemFees'` && `amount-description=='FBAPerUnitFulfillmentFee'`인 라인만 채택,
  금액은 `abs()` + 콤마를 점으로 치환 후 float 변환 (`7,50` → `7.50`, EU 로케일). 같은 주문에 라인이 여러 개면 합산.

## 5. 알려진 영구 불가 케이스 — 재조사 금지

- **UK**: EU 정산 리포트에 `Non-Amazon UK`로 태그된 라인이 **단 하나도 없음** (전체 리포트 grep으로 두 번 확인, marketplaceId를 UK로
  쿼리해도 DE와 동일한 리포트셋만 나옴 — 별도 리포트 스트림 없음). Finances API도 self-created MCF 주문 전체에 대해 원천적으로 데이터 없음
  (UK 특정 주문으로 직접 확인함). UK는 이 방법으론 절대 못 채움 — Amazon Support 이슈로 별도 제기해야 함.
- **`GCX-`가 아닌 주문ID** (예: `CONSUMER-...`): 애초에 매칭 대상이 아님.
- **PowerArc**: 아래 참고.
- **최근에 막 발송된 주문**: 정산이 아직 안 됐을 뿐 버그 아님. `_SettlementFeeCache`에 없으면 그냥 기다리면 자동으로 채워짐
  (수동으로 이 프로세스 반복해봤자 새 데이터가 나올 리 없음 — 이미 없는 걸 다시 찾는 것).

### PowerArc 특이사항 (2026-08-04 기준)

PowerArc는 **Spigen EU와 완전히 별개의 Seller 계정** (Account Health/재고 전부 별도). 시트의 일부 MCF 주문이 PowerArc 계정에서
생성되었다고 사용자가 확인함 — 즉 이 주문들의 정산 데이터는 PowerArc 전용 Seller Central에만 있고 Spigen EU 자격증명으로는 절대 못 찾음.

**막힌 지점**: PowerArc의 "Manage Your Apps" (`https://sellercentral.de/apps/manage`, PowerArc 계정으로 전환 후)를 확인한 결과
`ChannelReply`, `Tquens` 두 개만 승인되어 있고, 우리가 쓰는 SP-API 앱은 전혀 연결 안 되어 있음. **PowerArc용 SP-API 자격증명 자체가
아직 없음.** 이 스킬을 PowerArc에 쓰려면 먼저:
1. 기존에 여러 계정에서 재사용 가능한 "발행형" Spigen SP-API 앱이 있는지 확인 (있으면 PowerArc Seller Central에서 그 앱을 Authorize만
   하면 됨 — 간단).
2. 없으면 PowerArc 소유의 Amazon Developer Central에서 새 self-authorized 앱을 처음부터 등록해야 함 (LWA client id/secret, AWS IAM
   role 등 — 이번 세션에서 완료 못함, 사용자 확인/작업 필요).
3. 자격증명이 생기면 `~/.sp-api-config.json`에 `lwa_client_id_powerarc` / `lwa_client_secret_powerarc` / `refresh_token_powerarc`
   패턴으로 추가하고, `sp-api-proxy.py`의 `lwa_token()` cred 분기에도 `"powerarc"` 케이스를 추가해야 이 스킬에서 재사용 가능.
4. PowerArc/Germany의 정확한 marketplaceId도 이때 함께 확인.

이 자격증명이 없는 상태에서 이 스킬이 실행되면: PowerArc로 식별되는 주문(= EU/JP 양쪽 정산 리포트에서 못 찾은 나머지 blank 행 중,
사용자가 PowerArc 것이라고 알려준 것)은 "PowerArc 자격증명 없어서 스킵" 이라고 명확히 보고하고 나머지(EU/JP)는 정상 진행할 것 —
전체 작업을 막지 말 것.

## 6. 시트에 쓰기 — 반드시 GAS 1회성 함수 경유, Python에서 직접 API로 쓰지 말 것

1. `~/Desktop/GCX/GAS_Operations/MCF_Tracking/` 에 `_oneTimeXBackfill.js` 새 파일 생성. **함수명은 언더스코어로 시작하면 안 됨**
   (GAS 에디터의 Run 드롭다운에서 `_`로 시작하는 함수는 숨겨짐 — 이번 세션에 두 번 걸림). 매칭된 fee map을 JS 객체 리터럴로 인라인.
2. 스킵 로직은 `backfillMCFFeesRecent()`와 동일하게: 현재 값이 빈칸/`'RETRY'`/에러값일 때만 덮어씀, 그 외엔 절대 안 건드림.
3. `clasp push --force` (해당 디렉토리에서, `cd` 확인 필수 — 새 Bash 호출마다 cwd가 리셋될 수 있음).
4. Apps Script 에디터(`https://script.google.com/u/0/home/projects/1kDfEUVEEJ7TCA3HOMF6EYFTjbeZZKIeg_X84wCbLT1-tQqJI2ZlPUCxp/edit`)를
   브라우저에서 열고: 파일 트리에서 새 파일 클릭(풀 포인터 이벤트) → 함수 셀렉터가 자동으로 그 파일의 첫 함수를 선택함 → "Run" 버튼 클릭
   (풀 포인터 이벤트).
5. **실행 중에 절대 페이지를 navigate하지 말 것** — 도중에 다른 페이지로 이동하면 진행 중이던 실행 요청이 죽어버림 (이번 세션에 실제로
   발생). "Loading..." 상태로 10~20초 걸리는 건 정상 (특히 `AMZTK()`/`AMZTK_JP()` 커스텀함수 burst와 GAS 실행 큐를 다툴 때) — 그냥
   페이지 안 건드리고 기다릴 것.
6. Execution log에서 `"done — written: N, already filled (skipped): M, not in map: K"` 라인으로 결과 확인.
7. 검증: `_diag`류 읽기 전용 함수로 매칭 fee map 중 5~10개 샘플을 실제 시트 값과 대조해서 정확히 일치하는지 확인 후에만 완료 보고.
8. 1회성 파일 삭제 시도 (best-effort): 로컬 삭제 후 `clasp push`는 "already up to date"만 뜨고 원격에 반영 안 되는 경우가 잦음 — 안 되면
   에디터 파일 리스트에서 `[aria-label="File operations menu for {filename}."]` 요소 클릭 → 뜨는 메뉴에서 "Delete" 클릭. 이것도 안 되면
   그냥 남겨둠 (기능상 무해, 미관상 문제일 뿐 — 여기 시간 낭비하지 말 것).

## 7. 최종 보고 형식

작업 끝나면 다음을 표/목록으로 정리해서 보고:
- 마켓플레이스별: 대상 행 수 / 채운 행 수 / 아직 빈 행 수
- 아직 빈 행의 원인 분류: (a) 정산 대기 중(자동으로 해결될 것) (b) UK/PowerArc 등 구조적으로 이 방법으론 불가 (c) 개별 조사 필요
- 실제로 시트에 쓰기 전에 매칭 결과를 먼저 보여주고 확인받을지, 바로 쓸지는 사용자가 이미 "쓰라"고 명시한 경우가 아니면 먼저 물어볼 것.

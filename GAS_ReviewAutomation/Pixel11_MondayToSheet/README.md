# Pixel11_MondayToSheet (Pixel 11 — Monday → Sheet, 전체 새로고침)

Monday.com **Pixel 11 Case+CP** 보드(`18425190666`)를 매일 오후 5시(KST) 자동으로 조회해서,
시트 데이터 전체를 Monday 보드 최신 상태로 **교체**합니다. Monday에서 삭제된 항목은 시트에서도 함께 사라집니다.

[GlxZ8_MondayToSheet](../GlxZ8_MondayToSheet/)(Galaxy Z8)와 **같은 코드 구조·같은 시트 레이아웃**이며, 보드만 다릅니다.
두 시트의 헤더가 동일하므로 Z8용 Looker Studio 대시보드를 그대로 복제해 데이터 소스만 바꾸면 됩니다.

| 항목 | 값 |
|---|---|
| Apps Script | `12VOkxNXwDDZIxu7xrMese-XHoUUPUYs5CXSwmTh9gYHQcvxpDNZcJqzJ` (시트 바운드) |
| 대상 스프레드시트 | `1iPFNSVo6gQ6FkWUGB-DSL0v-3IGq8frh87XaCWQxaOU` — `[Case+CP] Pixel 11 Series 클레임/배드리뷰_26/8/18-26/11/18_GCX` |
| 시트 탭 | `해외&국내 리뷰+클레임 데이터` |
| Monday 보드 | `18425190666` (📌Pixel 11 Case+CP) |
| Looker Studio | [`293263b0-b719-4094-950e-ce63eb499639`](https://datastudio.google.com/u/0/reporting/293263b0-b719-4094-950e-ce63eb499639) — Pixel 11 대시보드 (Z8 대시보드 `f9eb2a13-…` 구조 동일; 2026-09-11 문구·링크 Pixel 11로 교체 완료) |

---

## 동작 방식

1. Monday 보드에서 전체 항목을 가져옴 (필요한 컬럼만 조회)
2. formula 타입 컬럼(`인입사유`, `국가`, `데이터 출처`)은 Monday API 특성상 최초 응답에 비어있을 수 있어 2차 조회로 보정
3. 헤더(1행)를 제외한 기존 시트 데이터를 지우고, Monday에서 가져온 전체 항목으로 다시 채움 (전체 replace)
   - V~Y열(`월별`, `경과 일`, `경과 주`, `경과 월`)은 1행에 있는 `ARRAYFORMULA`가 계산하는 파생 컬럼이라 스크립트가 건드리지 않음
4. 수동 실행(메뉴) 시에는 진행 상황과 결과 팝업이 뜨고, 자동(매일) 실행 시에는 조용히 실행됨 (Apps Script 실행 로그에서 확인)

> ⚠️ 시트에서 수동으로 직접 편집한 내용은 다음 동기화 때 Monday 쪽 값으로 덮어써집니다. 시트를 기준 데이터로 편집하지 마세요.

## Z8 코드와 다른 점 (단 하나)

Pixel 11 보드에는 Z8 보드의 **`사진 유무 (Clean)`** formula 컬럼(`formula_mm7184wc`)이 없습니다.
그래서 `사진 유무` 시트 컬럼은 status 컬럼 `color_mm0ffgz8`을 읽은 뒤 스크립트 안의 `_photoYN_()`가
Z8 formula와 동일한 치환(`yes→Y`, `no→N`, 그 외 그대로)을 적용합니다.
이를 위해 `COLUMN_MAP` 항목에 선택적 `transform(v)` 훅이 추가되었습니다.

## 컬럼 매핑 (`Code.js`의 `COLUMN_MAP`)

| 시트 헤더 | Monday 컬럼 | Monday 컬럼 ID |
|---|---|---|
| item_id (A열, 내부용) | item id | — |
| Order ID | Name (제목) | — |
| Created 날짜 | Created 날짜 | `date_mm0f80th` |
| Purchased 날짜 | Purchased 날짜 | `date_mm59ejfp` |
| ASIN | ASIN (text 타입) | `text_mm0f1q4h` |
| SKU | SKU | `lookup_mm0fv615` |
| 대분류 | 대분류 | `lookup_mm0ffq8f` |
| 인입사유 | 인입사유 | `formula_mm0g81mb` |
| 국가 | 국가 | `formula_mm25vbf0` |
| 기종명 | 기종명 | `lookup_mm0f6j81` |
| 모델명 | 모델명 | `lookup_mm0fn79` |
| 색상명 | 색상명 | `lookup_mm0fg6ja` |
| 클레임/리뷰 | 클레임/리뷰 | `color_mm0f7bwq` |
| 생산업체 | 생산업체 | `lookup_mm0feh3b` |
| 원산지 | 원산지 | `lookup_mm0fahcy` |
| 고객 대응 | 고객대응 | `color_mm0fjzar` |
| Review Link | Review Link | `link_mm0fkspz` |
| Zendesk Ticket | Zendesk Ticket | `integration_mm0fzmv0` (`https://spigenhelp.zendesk.com/tickets/{id}` 형식으로 정리) |
| 데이터 출처 | 데이터 출처 | `formula_mm5hrmzb` |
| 사진 유무 | 사진 유무 (status) + `_photoYN_` | `color_mm0ffgz8` |
| Review Ratings | Review Ratings | `text_mm0fn5c0` |

> 보드에 `ASIN`이라는 이름의 컬럼이 2개(text 타입 / board_relation 연결형) 있는데, 여기서는 **text 타입**(`text_mm0f1q4h`)을 사용합니다.
> Pixel 11 보드의 컬럼 ID는 Z8 보드와 전부 동일합니다 (같은 템플릿에서 복제됨, 2026-09-11 확인).

---

## 설정 / 배포

- 코드 배포: 이 폴더에서 `clasp push --force` (`.clasp.json`에 스크립트 ID 포함)
- **스크립트 속성** `MONDAY_API_KEY` 필요 (프로젝트 설정 → 스크립트 속성; 코드에 직접 넣지 않음)
- 최초 1회 `setupDailyTrigger` 실행 → 매일 17:00(KST) `syncMondayToSheet` 트리거 등록
- 즉시 실행: 시트 메뉴 **Monday.com → 지금 동기화 (전체 새로고침)**
- 진단: **Monday.com → 보드 컬럼 ID 목록 보기** (`_diag_columns` 탭), **항목 1개 전체 컬럼 덤프**

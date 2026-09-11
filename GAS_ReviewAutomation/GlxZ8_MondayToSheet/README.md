# GlxZ8_MondayToSheet (Galaxy Z8 — Monday → Sheet, 전체 새로고침)

Monday.com **Galaxy Z8 Case+CP** 보드(`18421346787`)를 매일 오후 5시(KST) 자동으로 조회해서,
시트 데이터 전체를 Monday 보드 최신 상태로 **교체**합니다. Monday에서 삭제된 항목은 시트에서도 함께 사라집니다.

**대상 스프레드시트:** `1ojfYyewbRL9hSZWTED-O_BeJ4T4DeP3aHDBOsKzAL7s` (탭: `해외&국내 리뷰+클레임 데이터`)
**Apps Script:** `1u_4_2QI9fEy8vSV9F0Nv9NlvoxtizPoWAVFMd6J3gwK2UpkTxlIfmmb3` (시트 바운드, `.clasp.json` 포함)
**Pixel 11 버전:** [Pixel11_MondayToSheet](../Pixel11_MondayToSheet/) — 같은 코드, 보드 `18425190666`
**Monday 보드:** `18421346787` (📌Galaxy Z8 Case+CP)

---

## 동작 방식

1. Monday 보드에서 전체 항목을 가져옴 (필요한 컬럼만 조회)
2. formula 타입 컬럼(`인입사유`, `국가`)은 Monday API 특성상 최초 응답에 비어있을 수 있어 2차 조회로 보정
3. 헤더를 제외한 기존 시트 데이터를 지우고, Monday에서 가져온 전체 항목으로 다시 채움 (전체 replace)
4. 수동 실행(메뉴) 시에는 진행 상황과 결과(성공/실패, 처리 행 수)를 보여주는 팝업 창이 함께 뜸. 자동(매일) 실행 시에는 팝업 없이 조용히 실행되며, 결과는 Apps Script 실행 로그에서 확인 가능

> ⚠️ Order ID(리뷰 텍스트) 등 시트에서 수동으로 직접 편집한 내용이 있다면, 다음 동기화 때 Monday 쪽 값으로 덮어써집니다. 시트를 기준 데이터로 편집하지 마세요.

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
| Zendesk Ticket | Zendesk Ticket | `integration_mm0fzmv0` (URL을 `https://spigenhelp.zendesk.com/tickets/{id}` 형식으로 정리) |
| 데이터 출처 | 데이터 출처 | `formula_mm5hrmzb` |
| 사진 유무 | 사진 유무 (Clean) | `formula_mm7184wc` |
| Review Ratings | Review Ratings | `text_mm0fn5c0` |

> 보드에 `ASIN`이라는 이름의 컬럼이 2개(text 타입 / board_relation 연결형) 있는데, 여기서는 **text 타입**(`text_mm0f1q4h`)을 사용합니다.

컬럼 추가/제거는 `Code.js` 상단의 `COLUMN_MAP` 배열만 수정하면 됩니다.

---

## 최초 설정 방법

1. 대상 스프레드시트 열기 → **확장 프로그램 → Apps Script**
2. 기본 `Code.gs` 내용을 지우고 이 폴더의 `Code.js` 내용을 붙여넣기
3. `appsscript.json`은 편집기 좌측 ⚙ 프로젝트 설정에서 "appsscript.json 매니페스트 파일을 편집기에 표시" 체크 후 내용 반영
4. **프로젝트 설정 → 스크립트 속성**에서 속성 추가:
   - 키: `MONDAY_API_KEY`
   - 값: Monday.com API 토큰 (코드에 직접 넣지 마세요)
5. 함수 선택 드롭다운에서 `setupDailyTrigger` 선택 후 ▶ 실행 (최초 1회, 권한 승인 필요)
   → 매일 오후 5시경(KST) 자동 실행되는 트리거가 등록됩니다
6. 바로 테스트하려면 시트를 새로고침한 뒤 메뉴 **Monday.com → 지금 동기화 (전체 새로고침)** 클릭
   → 진행 상황과 결과(성공/실패, 처리 행 수)를 보여주는 팝업 창이 뜹니다

> 이미 예전 버전(신규 항목만 append하던 버전)의 일일 트리거를 등록해둔 상태라면, 코드 업데이트 후 **`setupDailyTrigger`를 한 번 다시 실행**해서 트리거를 새 함수(`syncMondayToSheet`)로 갱신하세요. (기존 트리거는 자동으로 삭제되고 새로 등록됩니다)

---

## 참고

- Apps Script 트리거는 정확히 17:00이 아니라 17:00~17:15 사이 정도에 실행될 수 있습니다 (Google 스케줄러 특성).
- 자동(매일) 실행은 팝업 없이 조용히 진행되며, 실행 기록/오류는 Apps Script 편집기의 **실행 로그**(왼쪽 시계 아이콘)에서 확인 가능합니다.
- 수동 실행 시 뜨는 팝업은 실행 중에는 1초마다 진행 로그를 갱신해서 보여주고, 완료되면 성공/실패 여부와 처리된 행 수를 표시합니다.
- ⚠️ Monday API 토큰이 이번 작업 중 채팅에 노출되었으니, Monday.com에서 토큰을 재발급(rotate)하고 위 스크립트 속성 값을 새 토큰으로 교체하는 것을 권장합니다.

## clasp 연동 (선택)

다른 프로젝트처럼 clasp로 관리하려면, 3번 과정에서 생성된 Apps Script 프로젝트의 스크립트 ID를 확인 후:

```json
// .clasp.json
{
  "scriptId": "<여기에 스크립트 ID>",
  "rootDir": "."
}
```

파일을 이 폴더에 추가하면 `clasp push`로 이후 업데이트를 배포할 수 있습니다.

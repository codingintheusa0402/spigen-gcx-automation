# Vaultly — 가상 자산 시뮬레이션 뱅킹 앱 (iOS)

실제 은행을 흉내 내지 않는 **가상의 은행 브랜드**로 만든 뱅킹 시뮬레이터입니다.
모든 계좌·거래·이자·알림은 기기 안에서만 계산되는 모의(dry-run) 데이터입니다.

## 기능
- **계좌 설정**: 입출금 계좌 잔액을 원하는 금액으로 직접 설정 / 여러 계좌 관리
- **예약 입금 알림**: "엄마 → ₩5,000,000 용돈, 오후 4:00" 처럼 예약하면 그 시각에
  푸시 알림이 오고 잔액이 실제로 증가 (반복: 매일/매주/매월 지원, 앱이 꺼져 있어도
  다음 실행 시 자동 정산)
- **상품 (이자 실적용)**
  - 적금: 매월 자동이체 + 일 단위 이자 누적, 만기 시 이자(세후) 지급
  - 예금: 거치식, 일 단위 이자 누적
  - 대출: 원리금균등상환, 매월 이자·원금 자동 상환, 총자산에서 차감
- **재력 설정 (라이프스타일 시뮬레이션)**: 거지 / 서민 / 중산층 / 부자 / 재벌 중 선택하면
  그 생활에 맞는 입출금이 매일·매주·매월 자동 생성 (거지: 알바비→캐피탈 상환·연체이자로 소진,
  재벌: 투자사 수익배분·배당·채권 이자·임대료 유입). 지난 30일 내역 백필, 앱 미실행 기간도 정산
  - 급여는 연차마다 인상 + 주기적 승진, 지출은 매년 물가상승 반영
  - `EventLibrary` (132개 스펙, 티어별 1,000+ 변형): 보험료·공과금·구독·세금·의료·사고·과태료·
    경조사·보너스·배당·임대·부업·돌발 사건(분실, 사기, 상속, 소송…) 정기/1회성
  - 거지 엔딩: 잔액 바닥 → 캐피탈 자동 추가대출, 매월 이자·연체수수료 복리로 빚만 쌓임
- **시간 가속**: 실시간(1x) · 1시간=1일 · 1시간=1주 · 1일=1년 · 1달=100년(1,200x) (점프 없이 연속 진행).
  모든 엔진(월급·생활 흐름, 예약 입금, 이자·상환, 시세)이 가상 시각으로 진행
- **증권(모의 투자)**: 원화/달러 예수금, 보유 종목(국내/해외) 평가금·손익률, 수익률 정렬,
  현재가/평가금·$/원 토글, 모의 매수/매도, 시세 자동 변동(앱 내 랜덤워크) 또는 직접 설정
- **송금(모의)**: 내 계좌 간 / 외부 수취인 — 실제 거래 없음
- **테마**: TOSS(다크 리스트 레이아웃, `219,320원` 표기, 송금 pill, 채우기/보내기) · Nova(미니멀) · Heritage(클래식) · Orbit Pay(월렛) — 각각 라이트/다크 대응
- **금융사 이름 편집**: 은행·증권사·카드사·캐피탈 이름을 설정에서 직접 입력 (화면·거래내역·자동 생성 흐름에 모두 반영)
- **앱 잠금**: Face ID / 기기 암호 (설정 > 보안·알림)
- **서체**: Pretendard (국내 금융앱 표준 서체, OFL) 번들
- **화면 모드**: 시스템 / 라이트 / 다크 선택
- **무지출 챌린지 (가짜 소비)**: 사고 싶었지만 참은 것(배달·커피·택시·쇼핑…)을 기록하면 그 금액이
  '참은 돈' 가상 계좌에 입금. 연속 일수·이번 달 참은 돈·누적, 공유 카드(ImageRenderer), 업적
- **소비 리포트**: 월별 지출/수입, 카테고리 도넛(Swift Charts), 많이 쓴 곳 TOP 5, 일별 막대, 지난달 대비.
  자동 생성 거래는 `EventLibrary` 스펙 → `SpendingCategory` 태그, 나머지는 키워드 분류
- **저축 목표**: 목표별 전용 계좌 + 매월 자동이체(`ScheduledEvent.toAccountID` 내부 이체), 진행률, 채우기, 달성 업적
- **시나리오 온보딩**: 월급날 · 로또 당첨 · 재벌 2세 · 퇴사 후 1년 — 재력·잔액·예약 입금을 한 번에 설정 (설정에서도 가능)
- **인생 게임 (퀘스트 · 승급)**: 거지/서민으로 시작, 중산층·부자·재벌은 잠김. 며칠마다 현실적인 선택지가 옵니다
  (`QuestLibrary`, `scripts/gen_quests.py`로 생성, 320개: 이직·승진·보험·주거·투자·사업·가족·교육·건강·사치·유혹·불법 알바·세금).
  선택은 즉시 현금, 반복 흐름(보험료·부업·임대료), 연봉 배수·고용주, 지출/병원비 배수, 몇 달 뒤 결과(도박·사업), 승급/몰락으로 이어집니다.
  실손 없으면 병원비 1.8배, 이직은 상여금도 따라 오르지만 해고되면 실업급여, 불법 알바는 50% 벌금·징역.
  순자산 임계치(5천만/10억/500억)를 넘으면 승급 퀘스트가 나옵니다. 80세에 인생 결산. 알림 밖에서도 선택지가 액션 버튼으로 옵니다
- **12시간 = 인생**: `TimeSpeed.lifeIn12h`(43,800x)로 실제 12시간에 60년을 삽니다 (엔진 하루 루프 400일/틱)
- **전체 거래내역**: 1일·7일·30일·3개월·6개월·12개월 + 입금/출금 + 분류 칩 필터, 행 탭 → 거래 상세
- **실제 은행 푸시 형식**: `[은행명] 03/09 15:40 입금 31,440,000원 엄마 잔액 1,281,440원`
- **위젯 (`VaultlyWidget`)**: 홈 small/medium(잔액·총자산·다음 입금), 잠금화면 rectangular/circular/inline, 챌린지 위젯.
  App Group `group.com.kevinkim.vaultly`의 스냅샷 JSON으로 그림
- **Live Activity**: 4시간 이내 예약 입금을 잠금화면·Dynamic Island에 "입금 예정 16:00" 카운트다운으로 표시
- **알바 미니게임 + 오락**: 편의점·카페·택배 상하차·대리운전·서빙·전단지·과외·주식 단타에 더해
  **배달 스쿠터 레이스**(🍜 픽업 → 🏠 배달, 건당 3,500원 + 팁, 3번 넘어지면 응급실 5만원)와
  **블랙잭 (게임 머니)**(바이인의 10%씩 베팅, 딜러 17 스탠드, 3:2; 서민부터, 앱 안 만 18세 확인 후에만; 체력 소모 없음)
- **오늘의 일정 · 공휴일**: `KoreanHolidays`(고정 공휴일 + 음력 설날·추석·부처님오신날 + 대체공휴일) — 캐릭터는 공휴일에 쉬고,
  홈 캐릭터 카드의 📅로 그날의 시간표(30분 단위 병합)와 올해 공휴일·다음 공휴일 D-day를 봅니다
- **인생 이벤트** (`LifeEventEngine`): 교통사고(치료비, 실손 없으면 더 비쌈, 2주 부상 → 알바 불가), 경찰 단속(범칙금·과태료),
  구속(불법 선택으로 쌓인 `heat`가 있을 때만; 벌금 + 3~12개월 급여 중단), 실직(급여 3~6개월 중단 + 실업급여 4개월).
  하루 1회 판정, 30일 쿨다운. 내 인생 화면의 "인생 이벤트" 카드에 기록
- **되돌리기**: 마지막 선택을 결과 시트/선택의 기록에서 되돌리면 은행 전체가 선택 직전으로 복원되고 퀘스트가 다시 옵니다.
  스토어 빌드에서는 보상형 광고 시청 후 (Pro는 바로)
- **광고 (스토어 빌드만, `AdService`)**: AdMob 보상형(되돌리기) + 전면(알바 3회마다·선택 4회마다). ATT 추적 허용 요청·UMP 동의 포함.
  개인 빌드에는 광고 코드가 켜지지 않고, UI 테스트는 네트워크 없는 가짜 광고(`ad-overlay`)를 봅니다.
  Pro는 광고 제거 + 보너스 게임 머니 1,000만원. 출시 전 교체할 테스트 ID는 `AppStore/LISTING.md` 참고
- **스토어 변형 (`Vaultly-Store` 스킴, `-D STORE`)**: 가상 브랜드만 사용(TOSS 테마 → 미니멀 다크/미니멀뱅크),
  앱 이름 "Vaultly 시뮬레이터", StoreKit 2 Pro(평생 9,900원 / 월 1,900원, `Vaultly.storekit` 로컬 테스트)로
  테마·시나리오·24x 이상 가속·공유 카드 잠금. 개인 빌드(`Vaultly` 스킴)는 모든 기능 무료

## 빌드
```bash
brew install xcodegen
cd VaultlyBank
xcodegen generate
open Vaultly.xcodeproj
```
Xcode 15+, iOS 17+. 실기기(USB 연결) 설치는 `scripts/install_device.sh` 한 줄로 끝납니다 (개인 팀 자동 서명, 번들 ID `com.kevinkim.vaultly`).

테스트 (엔진·이자·시계·라이프스타일·거래):
```bash
xcodebuild test -project Vaultly.xcodeproj -scheme Vaultly -destination 'platform=iOS Simulator,name=iPhone 15'
```

### UI 테스트 런치 인자
`-uitest-reset` (저장 초기화) · `-uitest-theme toss` · `-uitest-tier working` · `-uitest-quest ins_01` ·
`-uitest-clock 2026-10-03T10:00:00` (가상 시각 고정) · `-uitest-lifeevent accident|police|jail|jobLoss` · `-uitest-pro` (스토어 빌드 Pro 흉내).
`-uitest-*`가 하나라도 있으면 광고는 가짜 오버레이로 대체됩니다.

## 구조
```
Vaultly/
  App/        앱 진입점, BankStore(상태·저장·정산 오케스트레이션)
  Models/     Account, Transaction, ScheduledEvent, ProductTerms, Settings
  Engine/     ScheduleEngine(예약 정산), InterestEngine(이자/상환 정산)
  Services/   Persistence(JSON 저장), NotificationService(로컬 알림), WidgetBridge, LiveActivityService, StoreService(StoreKit 2)
  Shared/     앱·위젯 공용 (WidgetSnapshot, DepositActivityAttributes)
  Theme/      테마 정의
  Views/      SwiftUI 화면
VaultlyWidget/  WidgetKit 확장 (잔액·챌린지 위젯, Live Activity)
```

## Quality gate (every change, before every push)
`bash scripts/qa.sh` regenerates the project and the fictional-name table, then builds and tests **both** builds on
the iOS Simulator:
- **Personal (`Vaultly`):** unit tests (engines, quests, mini games, money rules) plus every UI walkthrough: onboarding, all tabs, quests, challenges, goals, and every 알바 mini game at 거지 and 중산층.
- **App Store (`Vaultly-Store`, STORE flag):** the same UI walkthroughs on the store build.

All screenshots are exported to `build/qa/<run>/<scheme>/`, and a summary is written to `build/qa/latest/REPORT.md` (pass/fail counts and failures). Install the guard once per clone with `bash scripts/install_qa_hook.sh`. After that, any push containing Vaultly changes runs the gate first and is **refused if it fails**. It's also refused if the remote isn't the personal `codingintheusa0402` account, so Vaultly never reaches HQ_GCX. New features ship with UI tests for their screens, so the walkthrough always covers the whole app.

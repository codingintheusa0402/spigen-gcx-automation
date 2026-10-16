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
- **실제 은행 푸시 형식**: `[은행명] 03/09 15:40 입금 31,440,000원 엄마 잔액 1,281,440원`
- **위젯 (`VaultlyWidget`)**: 홈 small/medium(잔액·총자산·다음 입금), 잠금화면 rectangular/circular/inline, 챌린지 위젯.
  App Group `group.com.kevinkim.vaultly`의 스냅샷 JSON으로 그림
- **Live Activity**: 4시간 이내 예약 입금을 잠금화면·Dynamic Island에 "입금 예정 16:00" 카운트다운으로 표시
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

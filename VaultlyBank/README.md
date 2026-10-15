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
- **증권(모의 투자)**: 원화/달러 예수금, 보유 종목(국내/해외) 평가금·손익률, 수익률 정렬,
  현재가/평가금·$/원 토글, 모의 매수/매도, 시세 자동 변동(앱 내 랜덤워크) 또는 직접 설정
- **송금(모의)**: 내 계좌 간 / 외부 수취인 — 실제 거래 없음
- **테마**: Nova(미니멀) · Heritage(클래식) · Orbit Pay(월렛) — 모두 가상 브랜드
- 모든 화면에 **SIM(시뮬레이션)** 배지 표시 (App Store 1.1.6 대응)

## 빌드
```bash
brew install xcodegen
cd VaultlyBank
xcodegen generate
open Vaultly.xcodeproj
```
Xcode 15+, iOS 17+. Signing Team만 지정하면 실기기 설치가 가능합니다.

## 구조
```
Vaultly/
  App/        앱 진입점, BankStore(상태·저장·정산 오케스트레이션)
  Models/     Account, Transaction, ScheduledEvent, ProductTerms, Settings
  Engine/     ScheduleEngine(예약 정산), InterestEngine(이자/상환 정산)
  Services/   Persistence(JSON 저장), NotificationService(로컬 알림)
  Theme/      테마 정의
  Views/      SwiftUI 화면
```

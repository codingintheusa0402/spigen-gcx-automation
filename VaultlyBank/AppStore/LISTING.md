# Vaultly — App Store listing (Store build only)

Build the `Vaultly-Store` scheme (config `Store`, `-D STORE`). Never submit the personal scheme.

## Category & rating
- **Primary category:** Games → **Simulation**. Secondary: Games → Role Playing (or Casual). **Not Finance.**
- `LSApplicationCategoryType = public.app-category.simulation-games` is set in project.yml.
- **Age rating questionnaire (answer honestly):**
  - Simulated Gambling: **Frequent/Intense** — the 블랙잭 table (game money only, no cash-out, no real-money purchase of chips) and
    risky life choices settled by chance. This alone makes the app **17+** worldwide.
  - Mature/Suggestive Themes: **Infrequent/Mild** — choices mention illegal side jobs, fines and prison as consequences.
  - Everything else: None. Korea: expect **청소년이용불가 (18+)** from the GRAC self-rating because of the simulated card game;
    the app also gates the table behind an in-app "만 18세 이상입니다" confirmation (`AgeGateSheet`, stored in settings).
  - Contests/Unrestricted Web Access/Loot boxes: No. The card game never sells chips — chips come from the in-game salary only.
- Privacy (App Privacy section, because of AdMob):
  - **Data Used to Track You:** Identifiers → Device ID (only when the user allows the ATT prompt).
  - **Data Not Linked to You:** Identifiers (Device ID), Usage Data (Advertising Data, Product Interaction), Diagnostics (Crash Data),
    Location (Coarse Location — derived from IP by the ad network). Purpose: Third-Party Advertising, Analytics.
  - **Data Linked to You:** none. The app itself stores everything on-device; nothing is sent to Vaultly's own servers (there are none).
  - Privacy policy URL is required — host a one-page policy that names Google AdMob as the ad partner before submitting.

## Name, subtitle, keywords (draft)
- **Name:** Vaultly 시뮬레이터 (CFBundleDisplayName in the Store config)
- **Subtitle:** 거지에서 재벌까지, 인생 시뮬레이션 게임
- **Keywords:** 인생게임,시뮬레이션,재테크게임,경제게임,무지출챌린지,가상자산,부자되기,선택게임

## Description (draft)
가상의 돈으로 한 사람의 인생을 살아보는 시뮬레이션 게임입니다.
거지에서 시작해 이직·보험·투자·사업 같은 320가지 선택을 하며 서민, 중산층, 부자, 재벌로 올라가 보세요.
선택마다 미래의 자산 흐름이 달라지고, 후회할 선택도 있습니다. 최고 속도로 하면 12시간 만에 한 인생이 끝납니다.

- 모든 금액·계좌·회사는 가상입니다. 실제 은행·결제 서비스와 연결되지 않습니다.
- 앱 안의 회사·은행 이름은 모두 이 게임을 위해 만든 가상의 이름입니다.
- 무지출 챌린지: 참은 소비가 '참은 돈'으로 쌓입니다.

## Screenshots — rules
- Show the life game first: quest card, tier ladder, 인생 결산, 무지출 챌린지, 소비 리포트.
- The "모의 · 가상 자산" watermark and "모의" chips must be visible. No screenshot that reads as a real bank app.
- Never market sharing balance screenshots ("친구에게 자랑") — that is what 1.1.6 rejects.

## Notes for App Review (paste into App Review Information)
> Vaultly is a life-simulation game played with fictional money. The player starts poor and makes life decisions
> (job changes, insurance, investments, side businesses) that change a simulated balance; reaching thresholds unlocks
> higher wealth tiers. All banks, companies and merchants in the store build are fictional names. Every balance screen
> carries a "모의 · 가상 자산" (mock / virtual assets) watermark, push notifications are labelled "[모의]" / "[인생 게임]",
> and balances cannot be typed in manually — they only change through gameplay. The app has no network features,
> collects no data, and is not connected to any bank or payment service. Pro is a one-time in-app purchase (StoreKit 2).
> Ads: Google AdMob (rewarded ad to undo the last life choice, interstitials between mini games). Vaultly Pro removes ads.
> The 블랙잭 table uses in-game money only — chips cannot be bought, won money cannot be withdrawn — and sits behind an
> in-app 18+ confirmation; the app is rated 17+ for simulated gambling. Public holidays, random life events (accident,
> police fine, jail after illegal choices, job loss) and mini games are all part of the simulation.
> To see the core loop quickly: 설정 > 시뮬레이션 > 내 인생 > "최대 속도로 살기".

## Ads — AdMob (Store build only), tracking prompt, Pro
- SDK: `GoogleMobileAds` via Swift Package (project.yml `packages`), code in `Vaultly/Services/AdService.swift`, compiled only under `STORE`.
  The personal `Vaultly` scheme has **no ads at all**; Pro users in the Store build have no ads either.
- **Replace the TEST ids before submitting:** `VAULTLY_ADMOB_APP_ID` in project.yml (→ `GADApplicationIdentifier`) and
  `AdService.rewardedUnit` / `AdService.interstitialUnit` (both are Google's public test units today). Then register the app in
  AdMob and link it to the App Store record.
- Placements: **rewarded** = "다른 선택으로 되돌리기" on the quest outcome sheet and the latest 선택의 기록 entry (undo the last life
  choice; Pro undoes for free). **Interstitials** = after every 3rd 알바 payslip and every 4th quest answer. Never inside a game.
- **Tracking prompt (ATT)** is requested once after onboarding, before the first ad load; text = `NSUserTrackingUsageDescription`
  (허용하면 관심사에 맞는 광고… 거부해도 모든 기능 동일). Answer "Yes" to *Does this app use the Advertising Identifier (IDFA)?* →
  *Serve advertisements within the app*. `SKAdNetworkItems` carries Google's published list (50 ids).
- **UMP consent** (EEA/UK GDPR form) runs before the SDK starts; elsewhere it is a no-op. Set the GDPR message up in AdMob → Privacy & messaging.
- **Pro** (`com.kevinkim.vaultly.pro` lifetime 9,900원 / `.pro.monthly` 1,900원): removes every ad and credits a one-time
  **Pro 보너스 10,000,000원** of game money (`BankStore.grantProBonusIfNeeded`). Say so in the IAP description and in the paywall.
- UI tests never load the SDK: any `-uitest-*` launch argument switches `AdService` to an in-process fake ad (`ad-overlay`).

## What the Store build changes (code)
| Area | Store build |
|---|---|
| Brand names | `StoreNames.fix` (generated by `scripts/store_names.py`) in `Ledger.apply`, `Brand`, notifications; quests get an `#if STORE` copy |
| Transfer bank list | fictional banks only (`StoreNames.banks`) |
| Money pushes | `[모의 · <bank>] … (가상)`, title prefixed `[모의]` |
| Quest pushes | `[인생 게임] 새로운 선택 · …` |
| Interruption level | `.active` (personal build keeps `.timeSensitive`) |
| Watermark | `MockWatermark` over every screen, `MockChip` on 총자산, widget "모의" |
| Direct balance editing | hidden + guarded: 잔액 설정/채우기, hand-made scheduled deposits (증권/모의 투자 removed entirely 2026-10-18) |
| Ads | AdMob rewarded (되돌리기) + interstitials, ATT prompt, UMP consent; Pro = no ads + 보너스 게임 머니 |
| 블랙잭 | game-money table behind the in-app 18+ confirmation; chips are never sold |

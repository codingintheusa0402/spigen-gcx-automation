import Foundation

/// Build-time variant. `Vaultly-Store` scheme compiles with `-D STORE`:
///  - no real-brand defaults (the Toss-style theme becomes "미니멀 다크" / 미니멀뱅크)
///  - "시뮬레이터" wording in the app name and onboarding
///  - Pro gating for themes, scenarios and time acceleration
/// The personal build has none of this.
enum AppVariant {
    static let isStore: Bool = {
        #if STORE
        return true
        #else
        return false
        #endif
    }()
}

/// What Pro unlocks in the store variant. Everything is free in the personal build.
enum ProFeature: String {
    case noAds, bonusMoney, themes, scenarios, acceleration, widgets, challengeShare

    var title: String {
        switch self {
        case .noAds: return "광고 없음 (되돌리기도 광고 없이)"
        case .bonusMoney: return "보너스 게임 머니 1,000만원"
        case .themes: return "모든 테마"
        case .scenarios: return "재벌 2세 · 로또 시나리오"
        case .acceleration: return "24배 이상 시간 가속"
        case .widgets: return "홈·잠금화면 위젯"
        case .challengeShare: return "챌린지 공유 카드"
        }
    }
}

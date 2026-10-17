import Foundation

// MARK: - Part-time job mini games: shared core (pure Foundation — no SwiftUI, unit-tested on any platform)
//
// Every job is a small, deterministic state machine:  `tick(dt)` advances time, action methods are the player's
// taps, `isOver` ends the shift and `result` grades it. Views only draw state and forward taps, so the rules live
// here and are covered by MiniGameTests. Wages follow the 2026 Korean minimum wage (시간당 10,320원).

/// Tier as a plain level so the core doesn't depend on app models: 0 거지 · 1 서민 · 2 중산층 · 3 부자 · 4 재벌.
typealias TierLevel = Int

enum JobKind: String, Codable, CaseIterable, Identifiable {
    case convenience, cafe, parcel, designatedDriver, serving, flyers, tutoring, dayTrading, scooter, cards
    var id: String { rawValue }

    var title: String {
        switch self {
        case .convenience: return "편의점 알바"
        case .cafe: return "카페 알바"
        case .parcel: return "택배 상하차"
        case .designatedDriver: return "대리운전"
        case .serving: return "식당 서빙"
        case .flyers: return "전단지 알바"
        case .tutoring: return "과외"
        case .dayTrading: return "주식 단타"
        case .scooter: return "배달 스쿠터 레이스"
        case .cards: return "블랙잭 (게임 머니)"
        }
    }
    var emoji: String {
        switch self {
        case .convenience: return "🏪"
        case .cafe: return "☕️"
        case .parcel: return "📦"
        case .designatedDriver: return "🚗"
        case .serving: return "🍽️"
        case .flyers: return "📄"
        case .tutoring: return "📚"
        case .dayTrading: return "📈"
        case .scooter: return "🛵"
        case .cards: return "🃏"
        }
    }
    var pitch: String {
        switch self {
        case .convenience: return "바코드 찍고 거스름돈 정확히. 줄이 너무 길어지면 점장님 호출."
        case .cafe: return "주문은 점점 빨라진다. 아이스/핫, 사이즈, 샷 추가까지 헷갈리지 말 것."
        case .parcel: return "지역별 트럭에 박스를 테트리스처럼. 시간 안에 못 실으면 바닥에 떨어진다."
        case .designatedDriver: return "손님 차로 밤길 귀가. 부드럽게, 긁지 말고. 흠집은 내 돈으로 물어준다."
        case .serving: return "쟁반 균형 잡으면서 맞는 테이블로. 떨어뜨리면 끝."
        case .flyers: return "지나가는 사람에게 타이밍 맞춰 한 장씩. 바쁜 사람한텐 내밀지 말 것."
        case .tutoring: return "수학·영어·국어 빠른 퀴즈. 연속 정답이면 학부모 만족도 상승."
        case .dayTrading: return "1분 단타. 크게 벌 수도, 크게 잃을 수도 있다."
        case .scooter: return "🍜에서 픽업해 🏠에 배달. 식기 전에, 넘어지지 말고. 3번 넘어지면 응급실행."
        case .cards: return "하우스와 21 대결. 게임 머니로 바이인, 한 판에 바이인의 10%. 만 18세 이상."
        }
    }
    /// 오락: no 하루 체력 used, stake comes from the account, not a job.
    var isEntertainment: Bool { self == .cards }
    /// Who can take this job (by tier level) — 단타 needs capital; 과외 needs an education; 상하차/전단지 are for anyone who needs cash.
    var levels: ClosedRange<TierLevel> {
        switch self {
        case .convenience, .cafe, .serving, .flyers: return 0...2
        case .parcel, .designatedDriver: return 0...1
        case .tutoring: return 1...2
        case .dayTrading: return 2...4
        case .scooter: return 0...1
        case .cards: return 1...4
        }
    }
    func lockedReason(for level: TierLevel) -> String? {
        if levels.contains(level) { return nil }
        if level < levels.lowerBound {
            if self == .dayTrading { return "투자금이 필요해요 (중산층부터)" }
            if self == .cards { return "바이인할 돈이 필요해요 (서민부터)" }
            return "\(TierNames.name(levels.lowerBound)) 이상"
        }
        return "이 계층은 이런 알바를 안 해요"
    }
    /// Who pays (shown as the counterparty; App Store build fictionalizes brand names in Ledger.apply).
    var employer: String {
        switch self {
        case .convenience: return "GS25 알바"
        case .cafe: return "스타벅스 파트너"
        case .parcel: return "CJ대한통운 허브"
        case .designatedDriver: return "카카오T 대리"
        case .serving: return "김밥천국 홀서빙"
        case .flyers: return "전단지 일당"
        case .tutoring: return "과외비"
        case .dayTrading: return "주식 단타 손익"
        case .scooter: return "배민커넥트 정산"
        case .cards: return "하우스 (게임 머니)"
        }
    }
    /// What one shift pays at grade B (100%).
    var baseWage: Decimal {
        switch self {
        case .convenience: return MinimumWage.hourly * 4           // 4시간
        case .cafe: return 10_500 * 4
        case .parcel: return 15_000 * 4                            // 야간 상하차는 시급이 세다
        case .designatedDriver: return 32_000                      // 콜 1건 (수수료 20% 뗀 뒤)
        case .serving: return MinimumWage.hourly * 4
        case .flyers: return MinimumWage.hourly * 3
        case .tutoring: return 70_000                              // 2시간 수업
        case .dayTrading: return 0                                 // P&L instead of a wage
        case .scooter: return 0                                    // 건당 3,500원 + 팁 (ScooterGame.feePerDelivery)
        case .cards: return 0                                      // chips won or lost
        }
    }
    var shiftLabel: String {
        switch self {
        case .convenience, .cafe, .parcel, .serving: return "4시간"
        case .flyers: return "3시간"
        case .tutoring: return "2시간 수업"
        case .designatedDriver: return "콜 1건"
        case .dayTrading: return "1분 세션"
        case .scooter: return "60초 라이딩"
        case .cards: return "2분 세션"
        }
    }
}

enum MinimumWage { static let hourly: Decimal = 10_320 }   // 2026년 최저임금

enum TierNames {
    static let all = ["거지", "서민", "중산층", "부자", "재벌"]
    static func name(_ level: TierLevel) -> String { all.indices.contains(level) ? all[level] : "?" }
}

enum Grade: String, Codable, Comparable {
    case F, C, B, A, S
    var multiplier: Decimal {
        switch self { case .F: return 0.4; case .C: return 0.8; case .B: return 1.0; case .A: return 1.15; case .S: return 1.3 }
    }
    var comment: String {
        switch self {
        case .S: return "사장님이 정직원 제안할 기세"
        case .A: return "칭찬받았다"
        case .B: return "무난한 하루"
        case .C: return "한 소리 들었다"
        case .F: return "다음엔 부르지 않겠단다"
        }
    }
    private var rank: Int { ["F", "C", "B", "A", "S"].firstIndex(of: rawValue)! }
    static func < (a: Grade, b: Grade) -> Bool { a.rank < b.rank }

    /// Grade from a 0…1 performance ratio.
    static func from(_ r: Double) -> Grade {
        switch r {
        case 0.9...: return .S
        case 0.75..<0.9: return .A
        case 0.55..<0.75: return .B
        case 0.3..<0.55: return .C
        default: return .F
        }
    }
}

struct ShiftResult: Codable, Equatable {
    var kind: JobKind
    var score: Int
    var grade: Grade
    /// Money credited (negative = the shift cost money: 대리 damage, 단타 loss).
    var payout: Decimal
    var lines: [String]          // human summary shown on the result sheet
}

/// Rounds a won amount to 10원 like real payslips.
func roundWon(_ d: Decimal) -> Decimal {
    var x = d / 10, r = Decimal()
    NSDecimalRound(&r, &x, 0, .plain)
    return r * 10
}

func wage(_ kind: JobKind, _ grade: Grade, extra: Decimal = 0) -> Decimal {
    roundWon(kind.baseWage * grade.multiplier + extra)
}

/// Small deterministic RNG so games are reproducible in tests (SplitMix64).
struct GameRNG {
    private var s: UInt64
    init(seed: UInt64) { s = seed == 0 ? 0x9E3779B97F4A7C15 : seed }
    mutating func next() -> UInt64 {
        s &+= 0x9E3779B97F4A7C15
        var z = s
        z = (z ^ (z >> 30)) &* 0xBF58476D1CE4E5B9
        z = (z ^ (z >> 27)) &* 0x94D049BB133111EB
        return z ^ (z >> 31)
    }
    mutating func unit() -> Double { Double(next() >> 11) / Double(1 << 53) }
    mutating func range(_ lo: Double, _ hi: Double) -> Double { lo + (hi - lo) * unit() }
    mutating func int(_ r: ClosedRange<Int>) -> Int { r.lowerBound + Int(next() % UInt64(r.count)) }
    mutating func pick<T>(_ a: [T]) -> T { a[int(0...(a.count - 1))] }
    mutating func chance(_ p: Double) -> Bool { unit() < p }
    /// Standard normal (Box–Muller).
    mutating func normal() -> Double {
        let u1 = max(unit(), 1e-12), u2 = unit()
        return (-2 * log(u1)).squareRoot() * cos(2 * .pi * u2)
    }
}

/// Every job engine conforms; the view layer drives `tick` from a display link.
protocol ShiftGame {
    var kind: JobKind { get }
    var timeLeft: Double { get }
    var isOver: Bool { get }
    mutating func tick(_ dt: Double)
    func result() -> ShiftResult
}

// MARK: - Persistence: plays, bests, the daily shift limit

struct JobsState: Codable, Equatable {
    var plays: [String: Int] = [:]
    var best: [String: Int] = [:]
    var earned: [String: Decimal] = [:]
    var shiftDay: String = ""          // yyyy-MM-dd of the sim day the counter belongs to
    var shiftsToday: Int = 0
    static let shiftsPerDay = 3        // 하루 체력: 알바 최대 3탕

    func shiftsLeft(on day: String) -> Int { day == shiftDay ? max(0, Self.shiftsPerDay - shiftsToday) : Self.shiftsPerDay }

    mutating func record(_ r: ShiftResult, day: String) {
        if day != shiftDay { shiftDay = day; shiftsToday = 0 }
        if !r.kind.isEntertainment { shiftsToday += 1 }      // 오락 costs no 체력
        plays[r.kind.rawValue, default: 0] += 1
        best[r.kind.rawValue] = max(best[r.kind.rawValue] ?? 0, r.score)
        earned[r.kind.rawValue, default: 0] += r.payout
    }

    static func dayKey(_ date: Date, calendar: Calendar = .current) -> String {
        let c = calendar.dateComponents([.year, .month, .day], from: date)
        return String(format: "%04d-%02d-%02d", c.year ?? 0, c.month ?? 0, c.day ?? 0)
    }
}

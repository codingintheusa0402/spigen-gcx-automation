import Foundation

/// 주식 단타 — a 60-second session on one jumpy stock. Buy and sell as often as you like (each trade pays
/// commission, each sale pays transaction tax); whatever you hold is sold at the bell. Win or lose big —
/// but never more than the stake you brought (a 3× credit position is margin-called at −30%).
struct DayTradingGame: ShiftGame {
    static let names = ["에이치테크", "누리바이오", "한강로보틱스", "미래배터리", "스카이엔터", "대양조선"]
    static let commission = 0.00015, sellTax = 0.0018      // 매매 수수료 0.015% · 거래세 0.18%
    static let marginCall = -0.30

    let kind = JobKind.dayTrading
    let name: String
    let stake: Decimal
    let leverage: Double
    private(set) var timeLeft: Double = 60
    private(set) var price: Double
    private(set) var history: [Double]
    private(set) var cash: Double              // in won
    private(set) var shares: Double = 0
    private(set) var trades = 0
    private(set) var feesPaid: Double = 0
    private(set) var marginCalled = false
    private(set) var lastEvent: String?
    private var drift: Double = 0
    private var vol: Double = 0.010
    private var regimeIn: Double = 0
    private var stepAcc: Double = 0
    private var rng: GameRNG
    private let borrowed: Double

    init(seed: UInt64, stake: Decimal, leverage: Double = 1) {
        var r = GameRNG(seed: seed)
        name = r.pick(Self.names)
        let p0 = Double(r.int(8...120)) * 500
        price = p0
        history = [p0]
        self.stake = stake
        self.leverage = max(1, leverage)
        let s = NSDecimalNumber(decimal: stake).doubleValue
        borrowed = s * (self.leverage - 1)
        cash = s + borrowed
        rng = r
    }

    var equity: Double { cash + shares * price - borrowed }
    var stakeValue: Double { NSDecimalNumber(decimal: stake).doubleValue }
    var pnl: Double { equity - stakeValue }
    var holding: Bool { shares > 0 }
    var isOver: Bool { timeLeft <= 0 || marginCalled }

    mutating func tick(_ dt: Double) {
        guard !isOver else { return }
        timeLeft -= dt
        stepAcc += dt
        // One price print every 0.25 s: regime-switching random walk with occasional news spikes
        while stepAcc >= 0.25 {
            stepAcc -= 0.25
            regimeIn -= 0.25
            if regimeIn <= 0 {
                drift = rng.range(-0.004, 0.004)
                vol = rng.range(0.004, 0.016)
                regimeIn = rng.range(4, 10)
            }
            var ret = drift + vol * rng.normal()
            if rng.chance(0.006) {
                let up = rng.chance(0.5)
                ret += (up ? 1 : -1) * rng.range(0.04, 0.09)
                lastEvent = up ? "📰 호재 공시! 급등" : "📰 악재 뉴스! 급락"
            }
            ret = max(-0.29, min(0.29, ret))          // 상·하한가 30%
            price = max(100, (price * (1 + ret) / 5).rounded() * 5)   // 호가 단위
            history.append(price)
            if history.count > 240 { history.removeFirst(history.count - 240) }
        }
        if leverage > 1 && equity / stakeValue - 1 <= Self.marginCall {
            sellAll(reason: "반대매매! 증거금 부족으로 강제 청산")
            marginCalled = true
        }
        if timeLeft <= 0 && holding { sellAll(reason: "장 마감 — 보유 물량 자동 매도") }
    }

    /// Buy with all available cash (full position).
    mutating func buy() {
        guard !isOver, !holding, cash > price else { return }
        let fee = cash * Self.commission
        shares = ((cash - fee) / price).rounded(.down)
        let cost = shares * price
        cash -= cost + cost * Self.commission
        feesPaid += cost * Self.commission
        trades += 1
        lastEvent = "매수 \(Int(shares))주 @ \(Int(price))"
    }

    mutating func sell() { sellAll(reason: nil) }

    private mutating func sellAll(reason: String?) {
        guard holding else { return }
        let proceeds = shares * price
        let fee = proceeds * (Self.commission + Self.sellTax)
        cash += proceeds - fee
        feesPaid += fee
        lastEvent = reason ?? "매도 \(Int(shares))주 @ \(Int(price))"
        shares = 0
        trades += 1
    }

    func result() -> ShiftResult {
        // Loss can never exceed the stake
        let p = max(-stakeValue, pnl)
        let payout = roundWon(Decimal(p))
        let pct = p / stakeValue
        let grade: Grade = pct >= 0.05 ? .S : pct >= 0.02 ? .A : pct >= 0 ? .B : pct > -0.05 ? .C : .F
        var lines = ["\(name) · 매매 \(trades)회 · 수수료·세금 \(Int(feesPaid))원",
                     String(format: "수익률 %@%.2f%%", pct >= 0 ? "+" : "", pct * 100)]
        if marginCalled { lines.append("반대매매로 강제 청산됐어요") }
        return ShiftResult(kind: kind, score: Int(p / 1000), grade: grade, payout: payout, lines: lines)
    }

    /// Stake choices offered by tier level (won).
    static func stakes(for level: TierLevel) -> [Decimal] {
        switch level {
        case 2: return [1_000_000, 5_000_000]
        case 3: return [10_000_000, 50_000_000]
        case 4: return [100_000_000, 1_000_000_000]
        default: return []
        }
    }
    static func leverageAllowed(for level: TierLevel) -> Bool { level >= 3 }   // 신용 3배는 부자부터
}

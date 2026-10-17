import Foundation

/// 카페 알바 — orders arrive faster and faster; build each drink exactly (menu, hot/iced, size, extra shot).
struct CafeGame: ShiftGame {
    enum Drink: String, CaseIterable { case americano = "아메리카노", latte = "카페라떼", vanilla = "바닐라라떼",
                                       greenTea = "녹차라떼", iceTea = "아이스티", strawberry = "딸기스무디"
        var emoji: String {
            switch self { case .americano: return "☕️"; case .latte: return "🥛"; case .vanilla: return "🍦"
            case .greenTea: return "🍵"; case .iceTea: return "🧋"; case .strawberry: return "🍓" }
        }
        /// Some drinks only come one way
        var icedOnly: Bool { self == .iceTea || self == .strawberry }
        var canShot: Bool { self == .americano || self == .latte || self == .vanilla }
    }
    struct Spec: Equatable {
        var drink: Drink
        var iced: Bool
        var large: Bool
        var extraShot: Bool
        var label: String {
            "\(iced ? "아이스" : "따뜻한") \(drink.rawValue) \(large ? "그란데" : "톨")" + (extraShot ? " +샷" : "")
        }
    }
    struct Order: Equatable { let id: Int; let spec: Spec; var patience: Double; let maxPatience: Double }

    let kind = JobKind.cafe
    private(set) var timeLeft: Double = 90
    private(set) var orders: [Order] = []
    /// The cup on the counter being built.
    var cup = Spec(drink: .americano, iced: true, large: false, extraShot: false)
    private(set) var served = 0
    private(set) var mixups = 0
    private(set) var expired = 0
    private(set) var lastEvent: String?
    private var spawnIn: Double = 1
    private var nextID = 0
    private var elapsed: Double = 0
    private var rng: GameRNG
    static let maxOrders = 5, maxFails = 5

    init(seed: UInt64) { rng = GameRNG(seed: seed) }

    var isOver: Bool { timeLeft <= 0 || mixups + expired >= Self.maxFails }
    /// Gap between orders shrinks from 9 s to 3.5 s over the shift.
    var spawnGap: Double { max(3.5, 9 - elapsed / 90 * 5.5) }

    private mutating func spawn() {
        guard orders.count < Self.maxOrders else { return }
        let d = rng.pick(Drink.allCases)
        let spec = Spec(drink: d, iced: d.icedOnly ? true : rng.chance(0.6), large: rng.chance(0.4),
                        extraShot: d.canShot && rng.chance(0.3))
        let p = max(14, 26 - elapsed / 10)
        orders.append(Order(id: nextID, spec: spec, patience: p, maxPatience: p))
        nextID += 1
    }

    mutating func tick(_ dt: Double) {
        guard !isOver else { return }
        timeLeft -= dt; elapsed += dt; spawnIn -= dt
        if spawnIn <= 0 { spawn(); spawnIn = spawnGap * rng.range(0.85, 1.15) }
        for i in orders.indices { orders[i].patience -= dt }
        let late = orders.filter { $0.patience <= 0 }.count
        if late > 0 { expired += late; orders.removeAll { $0.patience <= 0 }; lastEvent = "주문 취소: 너무 오래 걸렸어요" }
    }

    // Building the cup
    mutating func choose(_ d: Drink) {
        cup.drink = d
        if d.icedOnly { cup.iced = true }
        if !d.canShot { cup.extraShot = false }
    }
    mutating func setIced(_ v: Bool) { if !cup.drink.icedOnly || v { cup.iced = v } }
    mutating func setLarge(_ v: Bool) { cup.large = v }
    mutating func toggleShot() { if cup.drink.canShot { cup.extraShot.toggle() } }

    /// Hand the cup to order `id`. A cup that doesn't match the order exactly is a mix-up (remade, order stays).
    mutating func serve(to id: Int) {
        guard !isOver, let i = orders.firstIndex(where: { $0.id == id }) else { return }
        if orders[i].spec == cup {
            served += 1
            orders.remove(at: i)
            lastEvent = "\(cup.label) 나왔습니다!"
        } else {
            mixups += 1
            lastEvent = "손님: \"이거 제가 시킨 거 아닌데요?\""
        }
    }

    func result() -> ShiftResult {
        let score = max(0, served * 120 - mixups * 60 - expired * 40)
        let r = Double(served) / 14 - Double(mixups) * 0.07 - Double(expired) * 0.05
        let grade = mixups + expired >= Self.maxFails ? min(Grade.from(r), .C) : Grade.from(r)
        return ShiftResult(kind: kind, score: score, grade: grade, payout: wage(kind, grade),
                           lines: ["음료 \(served)잔 제조", "잘못 만든 음료 \(mixups)잔 · 취소된 주문 \(expired)건"])
    }
}

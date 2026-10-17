import Foundation

/// 배달 스쿠터 레이스 — ride through traffic, pick up orders at the 🍜 stand and drop them at the 🏠 markers before
/// they go cold. Three lanes, cars and buses to dodge; three crashes is an accident that ends the shift at the ER.
struct ScooterGame: ShiftGame {
    enum MarkerKind: String { case car = "🚗", bus = "🚌", dog = "🐕", pickup = "🍜", drop = "🏠" }
    struct Marker: Equatable { let id: Int; let kind: MarkerKind; var x: Double; var y: Double; var drift: Double = 0 }

    let kind = JobKind.scooter
    static let feePerDelivery: Decimal = 3_500
    static let erBill: Decimal = 50_000
    static let maxCrashes = 3
    static let shiftSeconds: Double = 60
    /// Collision half-width across the road (lanes 0.66 apart). 0.5 leaves no safe strip between lanes.
    static let hitHalfWidth = 0.5

    private(set) var timeLeft = ScooterGame.shiftSeconds
    private(set) var x: Double = 0                   // −1 … 1 across the road, lanes at −0.66 / 0 / 0.66
    private(set) var vx: Double = 0
    private(set) var speed: Double = 36
    private(set) var distance: Double = 0
    private(set) var markers: [Marker] = []
    private(set) var carrying = 0                    // orders on the scooter (max 3)
    private(set) var delivered = 0
    private(set) var cold = 0                        // orders dropped after too long (0 pay)
    private(set) var crashes = 0
    private(set) var lastEvent: String?
    private(set) var carryingSince: Double = 0
    var steer: Double = 0
    var boost = false                                // hold to go faster (and riskier)
    private var spawnIn: Double = 0.8
    private var nextID = 0
    private var hitCooldown: Double = 0
    private var rng: GameRNG

    init(seed: UInt64) { rng = GameRNG(seed: seed) }

    var accident: Bool { crashes >= Self.maxCrashes }
    var isOver: Bool { timeLeft <= 0 || accident }
    var elapsed: Double { Self.shiftSeconds - timeLeft }

    mutating func tick(_ dt: Double) {
        guard !isOver else { return }
        timeLeft -= dt
        hitCooldown = max(0, hitCooldown - dt)
        let ax = steer * 4.5 - vx * 3.0
        vx += ax * dt
        x = min(1, max(-1, x + vx * dt))
        if abs(x) >= 1 { vx = 0 }
        speed = (boost ? 54 : 36) + elapsed * 0.35            // traffic gets faster through the shift
        let dy = speed * dt
        distance += dy
        for i in markers.indices {
            markers[i].y -= dy
            if markers[i].kind == .dog { markers[i].x = max(-0.9, min(0.9, markers[i].x + markers[i].drift * dt)) }   // dogs wander
        }
        markers.removeAll { $0.y < -6 }
        if carrying > 0 { carryingSince += dt }
        spawnIn -= dt
        if spawnIn <= 0 {
            let lane = rng.pick([-0.66, 0, 0.66])
            let kind: MarkerKind
            if carrying == 0 { kind = rng.chance(0.45) ? .pickup : rng.pick([.car, .car, .bus, .dog]) }
            else { kind = rng.chance(0.35) ? .drop : (rng.chance(0.15) ? .pickup : rng.pick([.car, .car, .bus, .dog])) }
            markers.append(Marker(id: nextID, kind: kind, x: lane, y: 150, drift: kind == .dog ? (rng.chance(0.5) ? 0.35 : -0.35) : 0)); nextID += 1
            if kind != .pickup, kind != .drop, rng.chance(0.3 + 0.25 * elapsed / Self.shiftSeconds) {   // more double blocks later
                let other = rng.pick([-0.66, 0, 0.66].filter { $0 != lane })
                markers.append(Marker(id: nextID, kind: .car, x: other, y: 150)); nextID += 1
            }
            spawnIn = (boost ? 0.7 : 1.0) * rng.range(0.75, 1.2) * max(0.6, 1 - elapsed / 120)
        }
        if hitCooldown == 0, let i = markers.firstIndex(where: { abs($0.y) < 2.5 && abs($0.x - x) < Self.hitHalfWidth }) {
            let m = markers[i]
            switch m.kind {
            case .pickup:
                if carrying < 3 { carrying += 1; if carrying == 1 { carryingSince = 0 }; lastEvent = "주문 픽업 🍜 (\(carrying)건 싣는 중)" }
                else { lastEvent = "더 못 실어요 — 배달부터!" }
                markers.remove(at: i)
            case .drop:
                if carrying > 0 {
                    carrying -= 1
                    if carryingSince > 25 { cold += 1; lastEvent = "식은 음식… 별점 1점 ⭐️" }
                    else { delivered += 1; lastEvent = "배달 완료 🏠 +\(Self.feePerDelivery)원" }
                    if carrying == 0 { carryingSince = 0 }
                } else { lastEvent = "실은 주문이 없어요" }
                markers.remove(at: i)
            default:
                crashes += 1
                lastEvent = accident ? "🚑 사고! 응급실로…" : "쾅! 넘어졌다 (\(crashes)/\(Self.maxCrashes))"
                hitCooldown = 1.2
                vx = 0
            }
        }
    }

    func result() -> ShiftResult {
        let expected = 6.0
        var ratio = Double(delivered) / expected - Double(crashes) * 0.08 - Double(cold) * 0.05
        if accident { ratio = min(ratio, 0.35) }
        let grade = Grade.from(max(0, ratio))
        let score = max(0, delivered * 1000 - crashes * 300 - cold * 200 + (accident ? 0 : Int(timeLeft * 10)))
        let fees = Self.feePerDelivery * Decimal(delivered)
        let tip: Decimal = grade >= .A ? roundWon(fees * Decimal(double: grade == .S ? 0.3 : 0.15)) : 0
        let payout = fees + tip - (accident ? Self.erBill : 0)
        var lines = ["배달 \(delivered)건 · 식은 음식 \(cold)건", "사고 \(crashes)회"]
        if tip > 0 { lines.append("팁 \(tip)원") }
        if accident { lines.append("응급실 치료비 \(Self.erBill)원은 내 돈으로") }
        return ShiftResult(kind: kind, score: score, grade: grade, payout: roundWon(payout), lines: lines)
    }
}

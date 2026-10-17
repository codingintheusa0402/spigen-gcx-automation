import Foundation

/// 대리운전 — drive the customer's car home at night. Steer between lanes, dodge parked cars and potholes,
/// and keep it smooth: hard jerks annoy the customer, scrapes are repaired out of your own pay.
struct DriverGame: ShiftGame {
    enum ObstacleKind: String { case car = "🚙", taxi = "🚕", cone = "🚧", pothole = "🕳️" }
    struct Obstacle: Equatable { let id: Int; let kind: ObstacleKind; let x: Double; var y: Double }  // y: metres ahead of the car

    let kind = JobKind.designatedDriver
    static let routeLength: Double = 2_400          // metres to the customer's apartment
    static let repairPerScrape: Decimal = 60_000
    static let maxScrapes = 3
    /// Half-width of the collision box across the road. Lanes are 0.66 apart, so anything under 0.33 left a
    /// safe strip between two lanes where nothing could ever hit you — 0.5 closes it.
    static let hitHalfWidth = 0.5

    private(set) var timeLeft: Double = 90           // the customer won't wait longer than this
    /// Car position across the road, −1 (left edge) … 1 (right edge). Lanes are centred at −0.66, 0, 0.66.
    private(set) var x: Double = 0
    private(set) var vx: Double = 0
    private(set) var speed: Double = 28              // m/s, rises a little as the road empties
    private(set) var distance: Double = 0
    private(set) var obstacles: [Obstacle] = []
    private(set) var scrapes = 0
    private(set) var discomfort: Double = 0          // accumulated jerk
    private(set) var lastEvent: String?
    /// −1 left, 0 straight, +1 right (held by the player).
    var steer: Double = 0
    private var spawnIn: Double = 1.2
    private var nextID = 0
    private var hitCooldown: Double = 0
    private var rng: GameRNG

    init(seed: UInt64) { rng = GameRNG(seed: seed) }

    var arrived: Bool { distance >= Self.routeLength }
    var isOver: Bool { arrived || timeLeft <= 0 || scrapes >= Self.maxScrapes }
    var progress: Double { min(1, distance / Self.routeLength) }

    mutating func tick(_ dt: Double) {
        guard !isOver else { return }
        timeLeft -= dt
        hitCooldown = max(0, hitCooldown - dt)
        // Steering: accelerate sideways toward the held direction, with damping (a car doesn't teleport)
        let ax = steer * 3.2 - vx * 2.4
        let oldVx = vx
        vx += ax * dt
        x = min(1, max(-1, x + vx * dt))
        if abs(x) >= 1 { vx = 0 }
        discomfort += abs(vx - oldVx) / max(dt, 1e-3) > 6 ? dt * 1.2 : 0   // harsh lateral jerk
        // Road
        speed = min(34, speed + dt * 0.08)
        let dy = speed * dt
        distance += dy
        for i in obstacles.indices { obstacles[i].y -= dy }
        obstacles.removeAll { $0.y < -6 }
        spawnIn -= dt
        if spawnIn <= 0 && distance < Self.routeLength - 60 {
            let lane = rng.pick([-0.66, 0, 0.66])
            obstacles.append(Obstacle(id: nextID, kind: rng.pick([.car, .car, .taxi, .cone, .pothole]), x: lane, y: 140))
            nextID += 1
            if rng.chance(0.25 + 0.3 * min(1, distance / Self.routeLength)) {   // two lanes blocked more often as you get closer
                let other = rng.pick([-0.66, 0, 0.66].filter { $0 != lane })
                obstacles.append(Obstacle(id: nextID, kind: .car, x: other, y: 140)); nextID += 1
            }
            spawnIn = max(0.9, 2.2 - distance / Self.routeLength * 1.1) * rng.range(0.8, 1.2)
        }
        // Collision: car occupies y ∈ [−2, 2] and |Δx| < 0.33
        if hitCooldown == 0, let hit = obstacles.first(where: { abs($0.y) < 2.5 && abs($0.x - x) < Self.hitHalfWidth }) {
            if hit.kind == .pothole {
                discomfort += 1.5
                lastEvent = "쿵! 손님이 깼어요"
            } else {
                scrapes += 1
                lastEvent = "끼익— 범퍼 긁힘 (수리비 \(DriverGame.repairPerScrape)원)"
            }
            hitCooldown = 1.0
        }
    }

    var comfort: Double { max(0, 1 - discomfort / 12) }

    func result() -> ShiftResult {
        let repair = Self.repairPerScrape * Decimal(scrapes)
        let score = max(0, Int(progress * 1000) + Int(comfort * 500) - scrapes * 300)
        var r = progress * (0.5 + 0.5 * comfort) - Double(scrapes) * 0.15
        if !arrived { r = min(r, 0.5) }
        let grade = Grade.from(r)
        // Fare only if the customer got home; tips for a smooth ride; repairs always come out of your pocket
        let fare: Decimal = arrived ? wage(kind, grade) : 0
        let payout = fare - repair
        var lines = [arrived ? "손님 집 앞 도착 ✅" : (scrapes >= Self.maxScrapes ? "손님이 화나서 직접 운전하겠대요" : "시간 초과 — 손님이 택시를 불렀어요"),
                     "승차감 \(Int(comfort * 100))점 · 흠집 \(scrapes)곳"]
        if scrapes > 0 { lines.append("수리비 \(repair)원은 내 돈으로 변상") }
        return ShiftResult(kind: kind, score: score, grade: grade, payout: payout, lines: lines)
    }
}

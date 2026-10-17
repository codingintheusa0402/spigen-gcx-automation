import Foundation

/// 식당 서빙 — pick dishes from the kitchen pass onto your tray and take each to the table that ordered it.
/// The tray tips more the more you carry; tap ◀ ▶ to keep it level. Tip it over and everything on it is lost.
struct ServingGame: ShiftGame {
    static let menu: [(name: String, emoji: String)] = [("김치찌개", "🍲"), ("돈까스", "🍛"), ("라면", "🍜"), ("김밥", "🍙"),
                                                        ("비빔밥", "🥗"), ("떡볶이", "🌶️"), ("제육볶음", "🥘"), ("만두", "🥟")]
    static let tableCount = 6, trayMax = 3, maxFails = 5

    struct Table: Equatable { let number: Int; var wants: Int?; var wait: Double }   // wants = menu index
    struct Dish: Equatable { let id: Int; let menu: Int }

    let kind = JobKind.serving
    private(set) var timeLeft: Double = 80
    private(set) var tables: [Table] = (1...ServingGame.tableCount).map { Table(number: $0, wants: nil, wait: 0) }
    private(set) var pass: [Dish] = []          // ready at the kitchen
    private(set) var tray: [Dish] = []
    /// −1 … 1; |tilt| ≥ 1 tips the tray over.
    private(set) var tilt: Double = 0
    private var tiltVel: Double = 0
    private(set) var served = 0
    private(set) var wrongTable = 0
    private(set) var droppedDishes = 0
    private(set) var angry = 0
    private(set) var lastEvent: String?
    private var orderIn: Double = 0.5
    private var nextID = 0
    private var rng: GameRNG

    init(seed: UInt64) { rng = GameRNG(seed: seed) }

    var fails: Int { wrongTable + angry + (droppedDishes + 1) / 2 }
    var isOver: Bool { timeLeft <= 0 || fails >= Self.maxFails }

    mutating func tick(_ dt: Double) {
        guard !isOver else { return }
        timeLeft -= dt
        // New orders: an empty table orders; the kitchen sends the dish a few seconds later
        orderIn -= dt
        if orderIn <= 0, let i = tables.indices.filter({ tables[$0].wants == nil }).randomElement(using: &rng) {
            let m = rng.int(0...(Self.menu.count - 1))
            tables[i].wants = m
            tables[i].wait = 30
            if pass.count < 4 { pass.append(Dish(id: nextID, menu: m)); nextID += 1 }
            orderIn = rng.range(3, 6)
        }
        for i in tables.indices where tables[i].wants != nil {
            tables[i].wait -= dt
            if tables[i].wait <= 0 { tables[i].wants = nil; angry += 1; lastEvent = "\(tables[i].number)번 테이블: \"음식 언제 나와요?!\" 그냥 나감" }
        }
        // Tray physics: random wobble grows with the load; the player's taps push back
        if tray.isEmpty {
            tilt *= max(0, 1 - dt * 3); tiltVel = 0
        } else {
            let wobble = rng.normal() * Double(tray.count) * 0.9
            tiltVel += (wobble + tilt * 0.6) * dt
            tiltVel *= max(0, 1 - dt * 0.8)
            tilt += tiltVel * dt
            if abs(tilt) >= 1 {
                droppedDishes += tray.count
                lastEvent = "와장창! 쟁반을 엎었어요 (\(tray.count)그릇)"
                tray.removeAll(); tilt = 0; tiltVel = 0
            }
        }
        // A dish nobody wants any more is cleared from the pass
        var keep: [Dish] = []
        for d in pass {
            let wanted = tables.contains { $0.wants == d.menu }
            if wanted || !rng.chance(dt * 0.2) { keep.append(d) }
        }
        pass = keep
    }

    /// Counter-balance: −1 pushes left, +1 pushes right.
    mutating func balance(_ dir: Double) {
        guard !tray.isEmpty else { return }
        tiltVel += dir * 0.55
    }

    mutating func pickUp(_ dishID: Int) {
        guard !isOver, tray.count < Self.trayMax, let i = pass.firstIndex(where: { $0.id == dishID }) else { return }
        tray.append(pass.remove(at: i))
        tiltVel += rng.range(-0.3, 0.3)
    }

    /// Serve the matching dish from the tray to table `number` (if any). Wrong food = mistake, the dish stays on the tray.
    mutating func serve(table number: Int) {
        guard !isOver, let ti = tables.firstIndex(where: { $0.number == number }), let want = tables[ti].wants, !tray.isEmpty else { return }
        if let di = tray.firstIndex(where: { $0.menu == want }) {
            tray.remove(at: di)
            tables[ti].wants = nil
            served += 1
            lastEvent = "\(number)번 테이블 \(Self.menu[want].name) 나왔습니다~"
        } else {
            wrongTable += 1
            lastEvent = "\(number)번 테이블: \"이거 저희 거 아닌데요\""
        }
    }

    func result() -> ShiftResult {
        let score = max(0, served * 110 - wrongTable * 50 - droppedDishes * 40 - angry * 40)
        let r = Double(served) / 15 - Double(wrongTable) * 0.06 - Double(droppedDishes) * 0.04 - Double(angry) * 0.05
        let grade = fails >= Self.maxFails ? min(Grade.from(r), .C) : Grade.from(r)
        // Tips: a little extra per served table at grade A/S
        let tip: Decimal = grade >= .A ? Decimal(served) * 500 : 0
        var lines = ["\(served)테이블 서빙", "잘못 간 음식 \(wrongTable) · 깨진 그릇 \(droppedDishes) · 화난 손님 \(angry)"]
        if tip > 0 { lines.append("팁 \(tip)원") }
        return ShiftResult(kind: kind, score: score, grade: grade, payout: wage(kind, grade, extra: tip), lines: lines)
    }
}

extension Array {
    /// randomElement with the game's own RNG (deterministic in tests).
    func randomElement(using rng: inout GameRNG) -> Element? { isEmpty ? nil : self[rng.int(0...(count - 1))] }
}

import Foundation

/// 편의점 알바 — scan every item, then hand back the exact change before the line gets too long.
struct ConvenienceGame: ShiftGame {
    struct Item: Equatable { let name: String; let emoji: String; let price: Int }
    struct Customer: Equatable {
        let id: Int
        var items: [Item]
        var scanned: [Bool]
        let paid: Int                  // the bill handed over
        var patience: Double           // seconds before they walk out
        var total: Int { items.reduce(0) { $0 + $1.price } }
        var change: Int { paid - total }
        var allScanned: Bool { !scanned.contains(false) }
    }
    enum Phase: Equatable { case scanning, change }

    static let catalog: [Item] = [
        Item(name: "삼각김밥", emoji: "🍙", price: 1_500), Item(name: "컵라면", emoji: "🍜", price: 1_800),
        Item(name: "바나나우유", emoji: "🥛", price: 1_700), Item(name: "콜라", emoji: "🥤", price: 2_200),
        Item(name: "도시락", emoji: "🍱", price: 4_900), Item(name: "생수", emoji: "💧", price: 1_000),
        Item(name: "과자", emoji: "🍪", price: 2_400), Item(name: "캔커피", emoji: "☕️", price: 1_400),
        Item(name: "샌드위치", emoji: "🥪", price: 3_300), Item(name: "아이스크림", emoji: "🍦", price: 2_000),
        Item(name: "핫바", emoji: "🌭", price: 2_800), Item(name: "젤리", emoji: "🍬", price: 1_200),
    ]
    /// Buttons for building the change, largest first.
    static let denominations = [50_000, 10_000, 5_000, 1_000, 500, 100]

    let kind = JobKind.convenience
    private(set) var timeLeft: Double = 75
    private(set) var queue: [Customer] = []      // queue[0] is at the counter
    private(set) var phase: Phase = .scanning
    private(set) var changeDrawer = 0             // change built so far
    private(set) var served = 0
    private(set) var mistakes = 0
    private(set) var walkouts = 0
    private(set) var lastEvent: String? = nil
    private var spawnIn: Double = 0
    private var nextID = 0
    private var rng: GameRNG
    static let maxQueue = 6, maxMistakes = 5

    init(seed: UInt64) {
        rng = GameRNG(seed: seed)
        spawn(); spawn()
    }

    var current: Customer? { queue.first }
    var isOver: Bool { timeLeft <= 0 || queue.count > Self.maxQueue || mistakes >= Self.maxMistakes }

    private mutating func spawn() {
        let n = rng.int(1...(served < 4 ? 2 : 4))
        let items = (0..<n).map { _ in rng.pick(Self.catalog) }
        let total = items.reduce(0) { $0 + $1.price }
        // People pay with the next bill up (sometimes exact change in coins = paid == total)
        let bills = [1_000, 5_000, 10_000, 50_000].filter { $0 >= total }
        let paid = rng.chance(0.15) ? total : (bills.isEmpty ? ((total / 10_000) + 1) * 10_000 : rng.pick(Array(bills.prefix(2))))
        queue.append(Customer(id: nextID, items: items, scanned: Array(repeating: false, count: n), paid: paid,
                              patience: max(18, 34 - Double(served) * 0.8)))
        nextID += 1
        spawnIn = max(3.5, 9 - Double(served) * 0.35) * rng.range(0.8, 1.2)
    }

    mutating func tick(_ dt: Double) {
        guard !isOver else { return }
        timeLeft -= dt
        spawnIn -= dt
        if spawnIn <= 0 { spawn() }
        // Only people waiting in line lose patience; the one at the counter is being served
        for i in queue.indices.dropFirst() {
            queue[i].patience -= dt
        }
        guard !queue.isEmpty else { spawn(); return }
        let before = queue.count
        queue = [queue[0]] + queue.dropFirst().filter { $0.patience > 0 }
        if queue.count < before { walkouts += before - queue.count; lastEvent = "기다리다 지친 손님이 나갔어요" }
    }

    /// Scan item `i` of the customer at the counter. Scanning the same item twice is a mistake (double charge).
    mutating func scan(_ i: Int) {
        guard !isOver, phase == .scanning, var c = current, c.items.indices.contains(i) else { return }
        if c.scanned[i] { mistakes += 1; lastEvent = "이중 결제! 같은 상품을 두 번 찍었어요"; return }
        c.scanned[i] = true
        queue[0] = c
        lastEvent = nil
        if c.allScanned { phase = .change; changeDrawer = 0 }
    }

    mutating func addChange(_ amount: Int) {
        guard !isOver, phase == .change else { return }
        changeDrawer += amount
    }

    mutating func clearChange() { changeDrawer = 0 }

    /// Hand the change over. Exact → next customer. Wrong → mistake, start counting again.
    mutating func giveChange() {
        guard !isOver, phase == .change, let c = current else { return }
        if changeDrawer == c.change {
            served += 1
            queue.removeFirst()
            phase = .scanning
            changeDrawer = 0
            lastEvent = "감사합니다~ 안녕히 가세요"
            if queue.isEmpty { spawn() }
        } else {
            mistakes += 1
            lastEvent = changeDrawer > c.change ? "거스름돈을 더 줬어요!" : "거스름돈이 모자라요!"
            changeDrawer = 0
        }
    }

    func result() -> ShiftResult {
        let score = max(0, served * 100 - mistakes * 40 - walkouts * 30)
        // ~12 customers in 75 s is an excellent shift
        let r = Double(served) / 12 - Double(mistakes) * 0.06 - Double(walkouts) * 0.04
        let grade = queue.count > Self.maxQueue || mistakes >= Self.maxMistakes ? min(Grade.from(r), .C) : Grade.from(r)
        var lines = ["손님 \(served)명 계산", "실수 \(mistakes)번 · 그냥 나간 손님 \(walkouts)명"]
        if queue.count > Self.maxQueue { lines.append("줄이 너무 길어져서 점장님이 대신 들어왔어요") }
        return ShiftResult(kind: kind, score: score, grade: grade, payout: wage(kind, grade), lines: lines)
    }
}

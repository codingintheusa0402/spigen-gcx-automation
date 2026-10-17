import Foundation

/// 택배 상하차 — boxes come down the belt labelled with a region; drop each into the right truck and column,
/// Tetris-style, before it falls off the belt. A truck that's packed ≥ 80% leaves and a fresh one backs in.
struct ParcelGame: ShiftGame {
    static let regions = ["서울", "경기", "부산"]
    static let width = 5, height = 6

    struct Box: Equatable { let id: Int; let w: Int; let h: Int; let region: Int }
    struct Truck: Equatable {
        let region: Int
        var cells: [[Bool]] = Array(repeating: Array(repeating: false, count: ParcelGame.width), count: ParcelGame.height) // [row][col], row 0 = floor
        var filled: Int { cells.reduce(0) { $0 + $1.filter { $0 }.count } }
        var fill: Double { Double(filled) / Double(ParcelGame.width * ParcelGame.height) }
        /// Lowest row where a w×h box placed at `col` rests (gravity), or nil if it would stick out of the top.
        func restRow(col: Int, w: Int, h: Int) -> Int? {
            guard col >= 0, col + w <= ParcelGame.width else { return nil }
            var top = 0
            for c in col..<(col + w) {
                for r in stride(from: ParcelGame.height - 1, through: 0, by: -1) where cells[r][c] { top = max(top, r + 1); break }
            }
            return top + h <= ParcelGame.height ? top : nil
        }
        mutating func place(col: Int, row: Int, w: Int, h: Int) {
            for r in row..<(row + h) { for c in col..<(col + w) { cells[r][c] = true } }
        }
    }

    let kind = JobKind.parcel
    private(set) var timeLeft: Double = 75
    private(set) var trucks: [Truck] = ParcelGame.regions.indices.map { Truck(region: $0) }
    private(set) var box: Box
    private(set) var nextBox: Box
    private(set) var boxTimer: Double          // seconds before the current box falls off the belt
    private(set) var placedCells = 0
    private(set) var departures = 0
    private(set) var misSorted = 0
    private(set) var dropped = 0
    private(set) var lastEvent: String?
    private var nextID = 0
    private var elapsed: Double = 0
    private var rng: GameRNG
    static let maxFails = 6

    init(seed: UInt64) {
        var r = GameRNG(seed: seed)
        func make(_ id: Int) -> Box {
            let shapes = [(1, 1), (2, 1), (1, 2), (2, 2), (3, 1)]
            let s = r.pick(shapes)
            return Box(id: id, w: s.0, h: s.1, region: r.int(0...2))
        }
        box = make(0); nextBox = make(1)
        rng = r; nextID = 2
        boxTimer = 7
    }

    var isOver: Bool { timeLeft <= 0 || misSorted + dropped >= Self.maxFails }
    /// Time allowed per box shrinks from 7 s to 2.5 s.
    var boxTime: Double { max(2.5, 7 - elapsed / 75 * 4.5) }

    private mutating func advance() {
        box = nextBox
        let shapes = [(1, 1), (2, 1), (1, 2), (2, 2), (3, 1)]
        let s = rng.pick(shapes)
        nextBox = Box(id: nextID, w: s.0, h: s.1, region: rng.int(0...2))
        nextID += 1
        boxTimer = boxTime
    }

    mutating func tick(_ dt: Double) {
        guard !isOver else { return }
        timeLeft -= dt; elapsed += dt; boxTimer -= dt
        if boxTimer <= 0 { dropped += 1; lastEvent = "박스가 벨트에서 떨어졌어요!"; advance() }
    }

    /// Drop the current box into truck `t` at column `col`.
    mutating func place(truck t: Int, col: Int) {
        guard !isOver, trucks.indices.contains(t) else { return }
        guard let row = trucks[t].restRow(col: col, w: box.w, h: box.h) else {
            lastEvent = "거기엔 안 들어가요"   // not a fail — the player can still choose another spot in time
            return
        }
        if trucks[t].region != box.region { misSorted += 1; lastEvent = "\(Self.regions[box.region]) 박스를 \(Self.regions[t]) 트럭에! 오배송" }
        else { lastEvent = nil }
        trucks[t].place(col: col, row: row, w: box.w, h: box.h)
        placedCells += box.w * box.h
        if trucks[t].fill >= 0.8 {
            departures += 1
            lastEvent = "\(Self.regions[t]) 트럭 출발! 🚚"
            trucks[t] = Truck(region: t)
        }
        advance()
    }

    func result() -> ShiftResult {
        let score = max(0, placedCells * 10 + departures * 200 - misSorted * 80 - dropped * 50)
        let r = Double(placedCells) / 70 + Double(departures) * 0.05 - Double(misSorted) * 0.08 - Double(dropped) * 0.05
        let grade = misSorted + dropped >= Self.maxFails ? min(Grade.from(r), .C) : Grade.from(r)
        return ShiftResult(kind: kind, score: score, grade: grade, payout: wage(kind, grade),
                           lines: ["트럭 \(departures)대 출발 · 실은 칸 \(placedCells)",
                                   "오배송 \(misSorted)개 · 떨어뜨린 박스 \(dropped)개"])
    }
}

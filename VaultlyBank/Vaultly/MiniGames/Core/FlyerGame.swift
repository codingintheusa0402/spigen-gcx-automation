import Foundation

/// 전단지 알바 — people walk past; tap "건네기" while someone is in your hand zone. Perfect timing gets taken,
/// sloppy timing gets ignored, and pushing a flyer at someone busy (on the phone, in a hurry) gets you a glare.
struct FlyerGame: ShiftGame {
    enum Mood: String { case open = "🙂", neutral = "😐", busy = "📱" }
    struct Walker: Equatable { let id: Int; let mood: Mood; var x: Double; let speed: Double; var done: Bool }

    let kind = JobKind.flyers
    static let quota = 30, zone = 0.5, zoneHalf = 0.09, perfectHalf = 0.03, maxGlares = 5
    private(set) var timeLeft: Double = 60
    private(set) var walkers: [Walker] = []
    private(set) var handed = 0
    private(set) var perfect = 0
    private(set) var ignored = 0
    private(set) var glares = 0
    private(set) var lastEvent: String?
    private var spawnIn: Double = 0.3
    private var nextID = 0
    private var rng: GameRNG

    init(seed: UInt64) { rng = GameRNG(seed: seed) }

    var isOver: Bool { timeLeft <= 0 || handed >= Self.quota || glares >= Self.maxGlares }

    mutating func tick(_ dt: Double) {
        guard !isOver else { return }
        timeLeft -= dt
        spawnIn -= dt
        if spawnIn <= 0 {
            let mood: Mood = rng.chance(0.25) ? .busy : (rng.chance(0.4) ? .open : .neutral)
            walkers.append(Walker(id: nextID, mood: mood, x: 0, speed: rng.range(0.22, 0.42) * (mood == .busy ? 1.4 : 1), done: false))
            nextID += 1
            spawnIn = rng.range(0.6, 1.4)
        }
        for i in walkers.indices { walkers[i].x += walkers[i].speed * dt }
        walkers.removeAll { $0.x > 1.05 }
    }

    /// Offer a flyer to whoever is nearest the hand zone.
    mutating func hand() {
        guard let i = walkers.indices.filter({ !walkers[$0].done }).min(by: { abs(walkers[$0].x - Self.zone) < abs(walkers[$1].x - Self.zone) }) else { return }
        hand(to: walkers[i].id)
    }

    /// Offer a flyer to a specific passer-by (the player taps the person).
    mutating func hand(to id: Int) {
        guard !isOver, let i = walkers.firstIndex(where: { $0.id == id && !$0.done }) else { return }
        let off = abs(walkers[i].x - Self.zone)
        walkers[i].done = true
        if off > Self.zoneHalf { ignored += 1; lastEvent = "허공에 내밀었어요"; return }
        switch walkers[i].mood {
        case .busy:
            glares += 1; lastEvent = "📱 통화 중인 사람한테… 째려봄"
        case .open:
            handed += 1; if off <= Self.perfectHalf { perfect += 1 }; lastEvent = off <= Self.perfectHalf ? "완벽! 🙏" : "받아갔어요"
        case .neutral:
            if off <= Self.perfectHalf || rng.chance(0.55) {
                handed += 1; if off <= Self.perfectHalf { perfect += 1 }; lastEvent = "받아갔어요"
            } else { ignored += 1; lastEvent = "그냥 지나갔어요" }
        }
    }

    func result() -> ShiftResult {
        let score = max(0, handed * 40 + perfect * 20 - glares * 60 - ignored * 5)
        let r = Double(handed) / Double(Self.quota) + Double(perfect) * 0.005 - Double(glares) * 0.06
        let grade = glares >= Self.maxGlares ? min(Grade.from(r), .C) : Grade.from(r)
        var lines = ["\(handed)/\(Self.quota)장 배포 (완벽 타이밍 \(perfect))", "무시당함 \(ignored) · 째려봄 \(glares)"]
        if handed >= Self.quota { lines.append("할당량 다 돌리고 일찍 퇴근!") }
        return ShiftResult(kind: kind, score: score, grade: grade, payout: wage(kind, grade), lines: lines)
    }
}

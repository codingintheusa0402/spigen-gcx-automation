import Foundation

/// 과외 — 12 quick multiple-choice questions (수학 · 영어 · 국어), 9 seconds each. Streaks please the parents.
struct TutoringGame: ShiftGame {
    struct Question: Equatable { let subject: String; let prompt: String; let options: [String]; let answer: Int }

    let kind = JobKind.tutoring
    static let count = 12, perQuestion: Double = 9

    private(set) var questions: [Question]
    private(set) var index = 0
    private(set) var questionTime: Double = TutoringGame.perQuestion
    private(set) var correct = 0
    private(set) var streak = 0
    private(set) var bestStreak = 0
    private(set) var lastEvent: String?
    private(set) var lastWasCorrect: Bool?
    var timeLeft: Double { index >= questions.count ? 0 : questionTime + Double(questions.count - index - 1) * Self.perQuestion }
    var isOver: Bool { index >= questions.count }
    var current: Question? { isOver ? nil : questions[index] }

    init(seed: UInt64) {
        var rng = GameRNG(seed: seed)
        var qs: [Question] = []
        for i in 0..<Self.count {
            switch i % 3 {
            case 0: qs.append(Self.math(&rng, hard: i >= 6))
            case 1: qs.append(Self.english(&rng))
            default: qs.append(Self.korean(&rng))
            }
        }
        questions = qs
    }

    mutating func tick(_ dt: Double) {
        guard !isOver else { return }
        questionTime -= dt
        if questionTime <= 0 { streak = 0; lastWasCorrect = false; lastEvent = "시간 초과!"; nextQuestion() }
    }

    mutating func answer(_ option: Int) {
        guard let q = current else { return }
        if option == q.answer {
            correct += 1; streak += 1; bestStreak = max(bestStreak, streak)
            lastWasCorrect = true
            lastEvent = streak >= 3 ? "🔥 \(streak)연속 정답!" : "정답!"
        } else {
            streak = 0; lastWasCorrect = false
            lastEvent = "오답 — 정답: \(q.options[q.answer])"
        }
        nextQuestion()
    }

    private mutating func nextQuestion() { index += 1; questionTime = Self.perQuestion }

    func result() -> ShiftResult {
        let score = correct * 100 + bestStreak * 30
        let r = Double(correct) / Double(Self.count) + Double(bestStreak) * 0.01
        let grade = Grade.from(r)
        // Parents pay a bonus for a great session
        let bonus: Decimal = grade == .S ? 20_000 : 0
        var lines = ["\(Self.count)문제 중 \(correct)문제 정답 · 최고 \(bestStreak)연속"]
        if bonus > 0 { lines.append("학부모님 감동 — 보너스 20,000원") }
        return ShiftResult(kind: kind, score: score, grade: grade, payout: wage(kind, grade, extra: bonus), lines: lines)
    }

    // MARK: Question generators

    private static func shuffled(_ correct: String, _ wrong: [String], _ rng: inout GameRNG) -> ([String], Int) {
        var opts = Array(Set(wrong.filter { $0 != correct })).sorted().prefix(3).map { $0 }
        let at = rng.int(0...opts.count)
        opts.insert(correct, at: at)
        return (opts, at)
    }

    static func math(_ rng: inout GameRNG, hard: Bool) -> Question {
        if hard && rng.chance(0.5) {
            // Solve a·x + b = c with integer x
            let a = rng.int(2...9), x = rng.int(2...12), b = rng.int(1...20)
            let (o, i) = shuffled("\(x)", ["\(x + 1)", "\(x - 1)", "\(x + 2)", "\(x * 2)"], &rng)
            return Question(subject: "수학", prompt: "\(a)x + \(b) = \(a * x + b) 일 때 x는?", options: o, answer: i)
        }
        let a = rng.int(12...99), b = rng.int(3...9)
        let ans = a * b
        let (o, i) = shuffled("\(ans)", ["\(ans + b)", "\(ans - b)", "\(ans + 10)", "\(ans - 10)"], &rng)
        return Question(subject: "수학", prompt: "\(a) × \(b) = ?", options: o, answer: i)
    }

    static let vocab: [(String, String)] = [
        ("abandon", "버리다"), ("accurate", "정확한"), ("achieve", "성취하다"), ("borrow", "빌리다"), ("decline", "거절하다"),
        ("efficient", "효율적인"), ("enormous", "거대한"), ("frequent", "빈번한"), ("generous", "관대한"), ("hesitate", "망설이다"),
        ("inevitable", "피할 수 없는"), ("journey", "여정"), ("maintain", "유지하다"), ("negotiate", "협상하다"), ("obvious", "명백한"),
        ("permanent", "영구적인"), ("reluctant", "꺼리는"), ("sufficient", "충분한"), ("temporary", "일시적인"), ("vivid", "생생한"),
    ]
    static func english(_ rng: inout GameRNG) -> Question {
        let pick = rng.pick(vocab)
        let wrong = (0..<6).map { _ in rng.pick(vocab).1 }
        let (o, i) = shuffled(pick.1, wrong, &rng)
        return Question(subject: "영어", prompt: "'\(pick.0)'의 뜻은?", options: o, answer: i)
    }

    /// 맞춤법: (correct, wrong forms)
    static let spelling: [(String, [String])] = [
        ("어떻게 해", ["어떡게 해", "어떻해", "어떠케 해"]), ("왠지", ["웬지", "왼지", "웨지"]), ("며칠", ["몇일", "몇칠", "면칠"]),
        ("안 돼", ["안 되", "안되요", "안 대"]), ("금세", ["금새", "금쇄", "끔세"]), ("설거지", ["설겆이", "설걷이", "설거쥐"]),
        ("오랜만에", ["오랫만에", "오랜마네", "오랫마네"]), ("굳이", ["구지", "궂이", "굿이"]), ("일부러", ["일부로", "일브러", "일뿌러"]),
        ("희한하다", ["희안하다", "히한하다", "희한타다"]), ("역할", ["역활", "역핧", "여칼"]), ("어이없다", ["어의없다", "어이업다", "어의업다"]),
    ]
    static func korean(_ rng: inout GameRNG) -> Question {
        let p = rng.pick(spelling)
        let (o, i) = shuffled(p.0, p.1, &rng)
        return Question(subject: "국어", prompt: "맞춤법이 바른 것은?", options: o, answer: i)
    }
}

import Foundation

/// 과외 — 12 quick multiple-choice questions (수학 · 영어 · 국어), 8 seconds each, harder in the second half.
/// Nothing repeats inside a session: every bank is drawn without replacement, and the arithmetic/algebra/percent/
/// sequence generators are seeded per question.
struct TutoringGame: ShiftGame {
    struct Question: Equatable { let subject: String; let prompt: String; let options: [String]; let answer: Int }

    let kind = JobKind.tutoring
    static let count = 12, perQuestion: Double = 8

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
        var bank = Bank(rng: &rng)
        var qs: [Question] = []
        for i in 0..<Self.count {
            let hard = i >= Self.count / 2
            switch i % 3 {
            case 0: qs.append(Self.math(&rng, hard: hard))
            case 1: qs.append(hard && rng.chance(0.5) ? bank.grammar(&rng) : bank.english(&rng))
            default: qs.append(hard && rng.chance(0.5) ? bank.idiom(&rng) : bank.korean(&rng))
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
        let bonus: Decimal = grade == .S ? 20_000 : 0
        var lines = ["\(Self.count)문제 중 \(correct)문제 정답 · 최고 \(bestStreak)연속"]
        if bonus > 0 { lines.append("학부모님 감동 — 보너스 20,000원") }
        return ShiftResult(kind: kind, score: score, grade: grade, payout: wage(kind, grade, extra: bonus), lines: lines)
    }

    // MARK: Question generators

    private static func shuffled(_ correct: String, _ wrong: [String], _ rng: inout GameRNG) -> ([String], Int) {
        var pool = Array(Set(wrong.filter { $0 != correct }))
        pool.sort()
        var opts: [String] = []
        while opts.count < 3, !pool.isEmpty { opts.append(pool.remove(at: rng.int(0...(pool.count - 1)))) }
        let at = rng.int(0...opts.count)
        opts.insert(correct, at: at)
        return (opts, at)
    }

    static func math(_ rng: inout GameRNG, hard: Bool) -> Question {
        let kind = hard ? rng.int(0...4) : rng.int(0...1)
        switch kind {
        case 0:                                                   // 곱셈
            let a = rng.int(12...99), b = rng.int(3...9), ans = a * b
            let (o, i) = shuffled("\(ans)", ["\(ans + b)", "\(ans - b)", "\(ans + 10)", "\(ans - 10)", "\(ans + 2 * b)"], &rng)
            return Question(subject: "수학", prompt: "\(a) × \(b) = ?", options: o, answer: i)
        case 1:                                                   // 뺄셈·덧셈 섞기
            let a = rng.int(120...999), b = rng.int(45...399)
            let plus = rng.chance(0.5)
            let ans = plus ? a + b : max(a, b) - min(a, b)
            let (o, i) = shuffled("\(ans)", ["\(ans + 10)", "\(ans - 10)", "\(ans + 1)", "\(ans - 100)", "\(ans + 100)"], &rng)
            return Question(subject: "수학", prompt: plus ? "\(a) + \(b) = ?" : "\(max(a, b)) − \(min(a, b)) = ?", options: o, answer: i)
        case 2:                                                   // 일차방정식
            let a = rng.int(2...9), x = rng.int(2...12), b = rng.int(1...20)
            let (o, i) = shuffled("\(x)", ["\(x + 1)", "\(x - 1)", "\(x + 2)", "\(x * 2)", "\(max(1, x - 2))"], &rng)
            return Question(subject: "수학", prompt: "\(a)x + \(b) = \(a * x + b) 일 때 x는?", options: o, answer: i)
        case 3:                                                   // 백분율
            let base = rng.pick([200, 250, 400, 500, 800, 1200, 1500, 2000, 2500, 3000]), pct = rng.pick([5, 10, 15, 20, 25, 30, 40, 60, 75])
            let ans = base * pct / 100
            let (o, i) = shuffled("\(ans)", ["\(ans + base / 20)", "\(max(1, ans - base / 20))", "\(ans * 2)", "\(ans / 2)", "\(ans + 10)"], &rng)
            return Question(subject: "수학", prompt: "\(base)의 \(pct)%는?", options: o, answer: i)
        default:                                                  // 수열
            let start = rng.int(1...9)
            if rng.chance(0.5) {
                let d = rng.int(3...11)
                let seq = (0..<4).map { start + d * $0 }
                let ans = start + d * 4
                let (o, i) = shuffled("\(ans)", ["\(ans + d)", "\(ans - 1)", "\(ans + 1)", "\(ans + d / 2 + 1)", "\(ans - d)"], &rng)
                return Question(subject: "수학", prompt: "\(seq.map(String.init).joined(separator: ", ")), ?", options: o, answer: i)
            } else {
                let r = rng.int(2...3)
                let seq = (0..<4).map { start * Int(pow(Double(r), Double($0))) }
                let ans = start * Int(pow(Double(r), 4))
                let (o, i) = shuffled("\(ans)", ["\(ans + start)", "\(ans - start)", "\(ans / r + 1)", "\(ans * 2)", "\(ans + r)"], &rng)
                return Question(subject: "수학", prompt: "\(seq.map(String.init).joined(separator: ", ")), ?", options: o, answer: i)
            }
        }
    }

    /// Per-session bank: each list is shuffled once and consumed front to back, so no prompt repeats.
    struct Bank {
        private var vocab: [(String, String)]
        private var spelling: [(String, [String])]
        private var grammar: [(String, String, [String])]
        private var idioms: [(String, String)]

        init(rng: inout GameRNG) {
            func shuffle<T>(_ a: [T], _ rng: inout GameRNG) -> [T] { var x = a; for i in stride(from: x.count - 1, to: 0, by: -1) { x.swapAt(i, rng.int(0...i)) }; return x }
            vocab = shuffle(TutoringGame.vocab, &rng); spelling = shuffle(TutoringGame.spelling, &rng)
            grammar = shuffle(TutoringGame.grammarBank, &rng); idioms = shuffle(TutoringGame.idiomBank, &rng)
        }
        private static func take<T>(_ list: inout [T], _ all: [T], _ rng: inout GameRNG) -> T {    // static: no overlapping self access
            if list.isEmpty { list = all; for i in stride(from: list.count - 1, to: 0, by: -1) { list.swapAt(i, rng.int(0...i)) } }
            return list.removeFirst()
        }
        mutating func english(_ rng: inout GameRNG) -> Question {
            let pick = Self.take(&vocab, TutoringGame.vocab, &rng)
            let wrong = (0..<8).map { _ in rng.pick(TutoringGame.vocab).1 }
            let (o, i) = shuffled(pick.1, wrong, &rng)
            return Question(subject: "영어", prompt: "'\(pick.0)'의 뜻은?", options: o, answer: i)
        }
        mutating func grammar(_ rng: inout GameRNG) -> Question {
            let g = Self.take(&grammar, TutoringGame.grammarBank, &rng)
            let (o, i) = shuffled(g.1, g.2, &rng)
            return Question(subject: "영어", prompt: g.0, options: o, answer: i)
        }
        mutating func korean(_ rng: inout GameRNG) -> Question {
            let p = Self.take(&spelling, TutoringGame.spelling, &rng)
            let (o, i) = shuffled(p.0, p.1, &rng)
            return Question(subject: "국어", prompt: "맞춤법이 바른 것은?", options: o, answer: i)
        }
        mutating func idiom(_ rng: inout GameRNG) -> Question {
            let p = Self.take(&idioms, TutoringGame.idiomBank, &rng)
            let wrong = (0..<8).map { _ in rng.pick(TutoringGame.idiomBank).1 }
            let (o, i) = shuffled(p.1, wrong, &rng)
            return Question(subject: "국어", prompt: "'\(p.0)'의 뜻은?", options: o, answer: i)
        }
    }

    static let vocab: [(String, String)] = [
        ("abandon", "버리다"), ("accurate", "정확한"), ("achieve", "성취하다"), ("borrow", "빌리다"), ("decline", "거절하다"),
        ("efficient", "효율적인"), ("enormous", "거대한"), ("frequent", "빈번한"), ("generous", "관대한"), ("hesitate", "망설이다"),
        ("inevitable", "피할 수 없는"), ("journey", "여정"), ("maintain", "유지하다"), ("negotiate", "협상하다"), ("obvious", "명백한"),
        ("permanent", "영구적인"), ("reluctant", "꺼리는"), ("sufficient", "충분한"), ("temporary", "일시적인"), ("vivid", "생생한"),
        ("ambiguous", "모호한"), ("anticipate", "예상하다"), ("benefit", "이익"), ("cautious", "조심스러운"), ("compromise", "타협"),
        ("consequence", "결과"), ("deliberate", "고의적인"), ("diligent", "근면한"), ("eliminate", "제거하다"), ("evidence", "증거"),
        ("fragile", "깨지기 쉬운"), ("gratitude", "감사"), ("harsh", "가혹한"), ("immense", "엄청난"), ("indifferent", "무관심한"),
        ("legitimate", "합법적인"), ("modest", "겸손한"), ("notorious", "악명 높은"), ("persuade", "설득하다"), ("precise", "정밀한"),
        ("profound", "심오한"), ("prosper", "번영하다"), ("restrict", "제한하다"), ("scarce", "부족한"), ("subtle", "미묘한"),
        ("thrive", "번창하다"), ("urgent", "긴급한"), ("vague", "막연한"), ("withdraw", "철회하다"), ("abundant", "풍부한"),
        ("adequate", "적절한"), ("conceal", "숨기다"), ("devote", "바치다"), ("exaggerate", "과장하다"), ("flexible", "유연한"),
        ("humble", "겸손한"), ("innovative", "혁신적인"), ("mandatory", "의무적인"), ("overwhelm", "압도하다"), ("spontaneous", "자발적인"),
    ]

    /// 영어 문법 빈칸: (prompt, correct, wrong)
    static let grammarBank: [(String, String, [String])] = [
        ("She ___ to school every day.", "goes", ["go", "going", "gone"]),
        ("I have ___ in Seoul since 2020.", "lived", ["live", "living", "lives"]),
        ("If it ___ tomorrow, we will stay home.", "rains", ["rain", "rained", "raining"]),
        ("He is taller ___ his brother.", "than", ["then", "as", "from"]),
        ("There ___ many people at the party.", "were", ["was", "is", "be"]),
        ("She has been working here ___ five years.", "for", ["since", "during", "from"]),
        ("This is the book ___ I told you about.", "which", ["who", "whom", "whose"]),
        ("I look forward to ___ you.", "seeing", ["see", "saw", "seen"]),
        ("Neither of them ___ ready.", "is", ["are", "were", "be"]),
        ("By next year, I ___ my degree.", "will have finished", ["finish", "finished", "am finishing"]),
        ("The report ___ by the team yesterday.", "was written", ["wrote", "is written", "writes"]),
        ("He suggested that she ___ a doctor.", "see", ["sees", "saw", "seeing"]),
        ("___ you mind opening the window?", "Would", ["Could", "Should", "Will"]),
        ("I wish I ___ taller.", "were", ["am", "be", "being"]),
        ("She made me ___ for an hour.", "wait", ["to wait", "waiting", "waited"]),
        ("It's time we ___ home.", "went", ["go", "going", "gone"]),
    ]

    /// 맞춤법: (correct, wrong forms)
    static let spelling: [(String, [String])] = [
        ("어떻게 해", ["어떡게 해", "어떻해", "어떠케 해"]), ("왠지", ["웬지", "왼지", "웨지"]), ("며칠", ["몇일", "몇칠", "면칠"]),
        ("안 돼", ["안 되", "안되요", "안 대"]), ("금세", ["금새", "금쇄", "끔세"]), ("설거지", ["설겆이", "설걷이", "설거쥐"]),
        ("오랜만에", ["오랫만에", "오랜마네", "오랫마네"]), ("굳이", ["구지", "궂이", "굿이"]), ("일부러", ["일부로", "일브러", "일뿌러"]),
        ("희한하다", ["희안하다", "히한하다", "희한타다"]), ("역할", ["역활", "역핧", "여칼"]), ("어이없다", ["어의없다", "어이업다", "어의업다"]),
        ("뵙겠습니다", ["봽겠습니다", "뵈겠습니다", "뵙겟습니다"]), ("깨끗이", ["깨끗히", "깨끄시", "깨끋이"]), ("낳다 (아이를)", ["낫다 (아이를)", "낮다 (아이를)", "낟다 (아이를)"]),
        ("낫다 (병이)", ["낳다 (병이)", "낮다 (병이)", "났다 (병이)"]), ("됐다", ["됬다", "됀다", "됫다"]), ("틀리다 (답이)", ["다르다 (답이)", "틀니다 (답이)", "틀이다 (답이)"]),
        ("웬일이야", ["왠일이야", "웬닐이야", "왼일이야"]), ("오랫동안", ["오랜동안", "오렛동안", "오랬동안"]), ("곰곰이", ["곰곰히", "곰곰니", "곰고미"]),
        ("치르다", ["치루다", "치르더", "치뤄다"]), ("바람 (소망)", ["바램 (소망)", "바라음", "바럼"]), ("삼가다", ["삼가하다", "삼까다", "삼가하라"]),
        ("가르치다 (교육)", ["가리키다 (교육)", "가르키다", "갈치다"]), ("가리키다 (방향)", ["가르치다 (방향)", "가르키다 (방향)", "가리치다"]), ("설레다", ["설레이다", "설래다", "설레히다"]),
        ("쉼표", ["쉼표우", "쉼퓨", "쉬임표"]), ("돼요", ["되요", "되여", "돼여"]), ("안녕히 계세요", ["안녕히 계새요", "안녕이 계세요", "안녕히 게세요"]),
        ("눈곱", ["눈꼽", "눈곰", "눈꼿"]), ("개수", ["갯수", "게수", "개쑤"]), ("초점", ["촛점", "초쩜", "쵸점"]),
        ("찌개", ["찌게", "찌깨", "찌게찌"]), ("육개장", ["육계장", "육게장", "유개장"]), ("떡볶이", ["떡볶기", "떡복이", "떡뽁이"]),
    ]

    /// 사자성어·어휘: (word, meaning)
    static let idiomBank: [(String, String)] = [
        ("일석이조", "한 가지 일로 두 가지 이익"), ("타산지석", "남의 잘못도 내 교훈으로 삼음"), ("우유부단", "결단을 못 내리고 망설임"),
        ("대기만성", "큰 그릇은 늦게 이루어짐"), ("동병상련", "같은 처지끼리 서로 가엾게 여김"), ("고진감래", "고생 끝에 즐거움이 옴"),
        ("새옹지마", "인생의 길흉화복은 예측할 수 없음"), ("자업자득", "자기가 한 일의 결과를 자기가 받음"), ("설상가상", "어려운 일이 겹쳐 일어남"),
        ("역지사지", "상대의 처지에서 생각함"), ("주마가편", "잘하는 사람을 더 격려함"), ("온고지신", "옛것을 익혀 새것을 앎"),
        ("사면초가", "사방이 적으로 둘러싸여 고립됨"), ("과유불급", "지나침은 모자람만 못함"), ("유비무환", "미리 준비하면 근심이 없음"),
        ("호가호위", "남의 권세를 빌려 위세를 부림"), ("각주구검", "융통성 없이 낡은 생각을 고집함"), ("청출어람", "제자가 스승보다 나음"),
        ("결초보은", "죽어서도 은혜를 갚음"), ("오리무중", "일의 갈피를 잡을 수 없음"), ("조삼모사", "잔꾀로 남을 속임"),
        ("환골탈태", "완전히 새롭게 바뀜"), ("감언이설", "달콤한 말로 남을 꾐"), ("일취월장", "나날이 발전함"),
    ]
}

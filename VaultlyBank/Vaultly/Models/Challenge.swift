import Foundation

// MARK: - 무지출 챌린지 (fake-consumption / resisted purchases)

/// What the user wanted to buy but didn't. The amount is "deposited" into a virtual
/// 참은 돈 account so the saving shows up like real money.
enum ResistKind: String, Codable, CaseIterable, Identifiable {
    case delivery, coffee, taxi, shopping, convenience, drink, game, other
    var id: String { rawValue }
    var emoji: String {
        switch self {
        case .delivery: return "🛵"
        case .coffee: return "☕️"
        case .taxi: return "🚕"
        case .shopping: return "🛍️"
        case .convenience: return "🏪"
        case .drink: return "🍻"
        case .game: return "🎮"
        case .other: return "✨"
        }
    }
    var title: String {
        switch self {
        case .delivery: return "배달"
        case .coffee: return "커피"
        case .taxi: return "택시"
        case .shopping: return "쇼핑"
        case .convenience: return "편의점"
        case .drink: return "술자리"
        case .game: return "현질"
        case .other: return "기타"
        }
    }
    var defaultAmount: Decimal {
        switch self {
        case .delivery: return 24_000
        case .coffee: return 5_500
        case .taxi: return 14_000
        case .shopping: return 59_000
        case .convenience: return 6_800
        case .drink: return 45_000
        case .game: return 11_000
        case .other: return 10_000
        }
    }
    var placeholder: String {
        switch self {
        case .delivery: return "마라탕 배달"
        case .coffee: return "아아 한 잔"
        case .taxi: return "퇴근 택시"
        case .shopping: return "무신사 후드"
        case .convenience: return "야식 라면"
        case .drink: return "2차 술값"
        case .game: return "배틀패스"
        case .other: return "참은 것"
        }
    }
}

struct ResistedPurchase: Identifiable, Codable, Equatable {
    var id = UUID()
    var date: Date
    var title: String
    var amount: Decimal
    var kind: ResistKind
}

struct ChallengeState: Codable, Equatable {
    var items: [ResistedPurchase] = []
    var accountID: UUID? = nil          // 참은 돈 account that receives the virtual deposits
    var startedAt: Date? = nil

    var isActive: Bool { accountID != nil }

    func items(inMonthOf date: Date, calendar: Calendar = .current) -> [ResistedPurchase] {
        items.filter { calendar.isDate($0.date, equalTo: date, toGranularity: .month) }
    }

    func total(inMonthOf date: Date, calendar: Calendar = .current) -> Decimal {
        items(inMonthOf: date, calendar: calendar).reduce(0) { $0 + $1.amount }
    }

    var totalSaved: Decimal { items.reduce(0) { $0 + $1.amount } }

    /// Consecutive days (ending today or yesterday) with at least one resisted purchase.
    func streak(now: Date, calendar: Calendar = .current) -> Int {
        let days = Set(items.map { calendar.startOfDay(for: $0.date) })
        guard !days.isEmpty else { return 0 }
        var day = calendar.startOfDay(for: now)
        if !days.contains(day) {
            guard let y = calendar.date(byAdding: .day, value: -1, to: day), days.contains(y) else { return 0 }
            day = y
        }
        var n = 0
        while days.contains(day), let prev = calendar.date(byAdding: .day, value: -1, to: day) {
            n += 1; day = prev
        }
        return n
    }

    func bestStreak(calendar: Calendar = .current) -> Int {
        let days = Set(items.map { calendar.startOfDay(for: $0.date) }).sorted()
        var best = 0, run = 0
        var prev: Date?
        for d in days {
            if let p = prev, let next = calendar.date(byAdding: .day, value: 1, to: p), calendar.isDate(next, inSameDayAs: d) { run += 1 } else { run = 1 }
            best = max(best, run); prev = d
        }
        return best
    }
}

// MARK: - Savings goals

struct SavingsGoal: Identifiable, Codable, Equatable {
    var id = UUID()
    var title: String
    var emoji: String
    var target: Decimal
    var accountID: UUID            // account whose balance counts toward the goal
    var deadline: Date? = nil
    var autoEventID: UUID? = nil   // recurring internal transfer feeding the goal
    var createdAt: Date
    var reachedAt: Date? = nil

    func progress(balance: Decimal) -> Double {
        guard target > 0 else { return 0 }
        return min(1, max(0, (balance / target).doubleValue))
    }
}

// MARK: - Achievements

enum AchievementID: String, CaseIterable, Identifiable {
    case firstResist, streak7, streak30, saved100k, saved1m, firstGoal, firstProduct, millionaire, billionaire, timeTraveler   // firstTrade retired with 증권
    var id: String { rawValue }
    var title: String {
        switch self {
        case .firstResist: return "첫 참기"
        case .streak7: return "7일 연속"
        case .streak30: return "한 달 무지출"
        case .saved100k: return "10만 원 참음"
        case .saved1m: return "100만 원 참음"
        case .firstGoal: return "목표 달성"
        case .firstProduct: return "첫 금융상품"
        case .millionaire: return "천만장자"
        case .billionaire: return "억만장자"
        case .timeTraveler: return "시간 여행자"
        }
    }
    var detail: String {
        switch self {
        case .firstResist: return "무지출 챌린지에 첫 기록을 남겼어요"
        case .streak7: return "7일 연속으로 참았어요"
        case .streak30: return "30일 연속으로 참았어요"
        case .saved100k: return "참은 돈이 100,000원을 넘었어요"
        case .saved1m: return "참은 돈이 1,000,000원을 넘었어요"
        case .firstGoal: return "저축 목표를 처음 달성했어요"
        case .firstProduct: return "적금·예금·대출 상품에 처음 가입했어요"
        case .millionaire: return "총자산 10,000,000원 돌파"
        case .billionaire: return "총자산 1,000,000,000원 돌파"
        case .timeTraveler: return "시간 흐름을 24배 이상으로 올렸어요"
        }
    }
    var symbol: String {
        switch self {
        case .firstResist: return "hand.raised.fill"
        case .streak7, .streak30: return "flame.fill"
        case .saved100k, .saved1m: return "banknote.fill"
        case .firstGoal: return "target"
        case .firstProduct: return "doc.text.fill"
        case .millionaire, .billionaire: return "crown.fill"
        case .timeTraveler: return "clock.arrow.2.circlepath"
        }
    }
}

enum AchievementEngine {
    /// Returns achievements newly unlocked for `data` and records them.
    static func evaluate(now: Date, data: inout BankData, totalAssets: Decimal) -> [AchievementID] {
        let c = data.challenge
        var hit: [AchievementID] = []
        func check(_ id: AchievementID, _ cond: @autoclosure () -> Bool) {
            guard data.achievements[id.rawValue] == nil, cond() else { return }
            data.achievements[id.rawValue] = now
            hit.append(id)
        }
        check(.firstResist, !c.items.isEmpty)
        check(.streak7, c.bestStreak() >= 7)
        check(.streak30, c.bestStreak() >= 30)
        check(.saved100k, c.totalSaved >= 100_000)
        check(.saved1m, c.totalSaved >= 1_000_000)
        check(.firstGoal, data.goals.contains { $0.reachedAt != nil })
        check(.firstProduct, data.accounts.contains { $0.kind.isProduct && $0.terms != nil })
        check(.millionaire, totalAssets >= 10_000_000)
        check(.billionaire, totalAssets >= 1_000_000_000)
        check(.timeTraveler, data.clock.speed >= 24)
        return hit
    }
}

// MARK: - Scenarios (onboarding presets)

struct ScenarioEvent {
    let counterparty: String
    let memo: String
    let amount: Decimal
    let minutesFromNow: Int
    let repeatRule: RepeatRule
}

enum Scenario: String, CaseIterable, Identifiable {
    case payday, lottery, heir, jobless
    var id: String { rawValue }

    var emoji: String {
        switch self {
        case .payday: return "💼"
        case .lottery: return "🎰"
        case .heir: return "👑"
        case .jobless: return "📦"
        }
    }
    var title: String {
        switch self {
        case .payday: return "월급날"
        case .lottery: return "로또 당첨"
        case .heir: return "재벌 2세"
        case .jobless: return "퇴사 후 1년"
        }
    }
    var summary: String {
        switch self {
        case .payday: return "평범한 직장인. 10분 뒤 월급이 들어오고, 카드값·월세가 자동으로 빠져나갑니다."
        case .lottery: return "중산층으로 살다가 5분 뒤 1등 당첨금 2,030,000,000원이 입금됩니다."
        case .heir: return "재벌가 계좌. 임대료·배당·법인 급여가 매일 수억씩 들어옵니다."
        case .jobless: return "퇴직금 300만 원으로 시작. 알바와 캐피탈 대출로 버팁니다."
        }
    }
    var tier: WealthTier {
        switch self {
        case .payday: return .working
        case .lottery: return .middle
        case .heir: return .tycoon
        case .jobless: return .broke
        }
    }
    var startingBalance: Decimal? {
        switch self {
        case .jobless: return 3_000_000
        default: return nil
        }
    }
    var events: [ScenarioEvent] {
        switch self {
        case .payday: return [
            ScenarioEvent(counterparty: "(주)대성물류", memo: "급여", amount: 2_800_000, minutesFromNow: 10, repeatRule: .monthly),
            ScenarioEvent(counterparty: "{card}", memo: "카드 대금", amount: -1_180_000, minutesFromNow: 60 * 24 * 3, repeatRule: .monthly),
            ScenarioEvent(counterparty: "집주인", memo: "월세", amount: -650_000, minutesFromNow: 60 * 24 * 5, repeatRule: .monthly),
        ]
        case .lottery: return [
            ScenarioEvent(counterparty: "동행복권", memo: "로또 1등 당첨금 (세후)", amount: 2_030_000_000, minutesFromNow: 5, repeatRule: .none),
        ]
        case .heir: return [
            ScenarioEvent(counterparty: "VX홀딩스", memo: "배당금", amount: 850_000_000, minutesFromNow: 15, repeatRule: .weekly),
        ]
        case .jobless: return [
            ScenarioEvent(counterparty: "쿠팡 물류센터", memo: "일당", amount: 112_000, minutesFromNow: 30, repeatRule: .daily),
        ]
        }
    }
}

// MARK: - Spending categories (소비 리포트)

enum SpendingCategory: String, Codable, CaseIterable, Identifiable {
    case food, cafe, mart, shopping, transport, housing, medical, insurance, subscription, leisure, education, finance, tax, family, transfer, income, other
    var id: String { rawValue }
    var title: String {
        switch self {
        case .food: return "식비"
        case .cafe: return "카페·간식"
        case .mart: return "편의점·마트"
        case .shopping: return "쇼핑"
        case .transport: return "교통·차량"
        case .housing: return "주거·통신"
        case .medical: return "의료·건강"
        case .insurance: return "보험"
        case .subscription: return "구독"
        case .leisure: return "문화·여가"
        case .education: return "교육·육아"
        case .finance: return "금융·상환"
        case .tax: return "세금·공과금"
        case .family: return "경조사·가족"
        case .transfer: return "이체"
        case .income: return "수입"
        case .other: return "기타"
        }
    }
    var symbol: String {
        switch self {
        case .food: return "fork.knife"
        case .cafe: return "cup.and.saucer.fill"
        case .mart: return "cart.fill"
        case .shopping: return "bag.fill"
        case .transport: return "car.fill"
        case .housing: return "house.fill"
        case .medical: return "cross.case.fill"
        case .insurance: return "shield.fill"
        case .subscription: return "play.rectangle.fill"
        case .leisure: return "ticket.fill"
        case .education: return "book.fill"
        case .finance: return "building.columns.fill"
        case .tax: return "doc.text.fill"
        case .family: return "gift.fill"
        case .transfer: return "arrow.left.arrow.right"
        case .income: return "arrow.down.circle.fill"
        case .other: return "ellipsis.circle.fill"
        }
    }

    /// Category for an EventLibrary spec id.
    static func forSpec(_ id: String) -> SpendingCategory {
        switch id {
        case "dining", "working_ot_meal", "dining_fine": return .food
        case "convenience", "grocery": return .mart
        case "shopping", "shopping_lux", "electronics", "beauty": return .shopping
        case "transport", "taxi_premium", "car_fuel", "car_toll", "car_repair", "jet", "yacht": return .transport
        case "mgmt_fee", "util_elec", "util_gas", "util_water", "util_net", "util_phone", "repair_home", "laundry", "household_help", "moving", "deposit_pay": return .housing
        case "sub_ott", "sub_music", "sub_cloud", "sub_gym", "sub_app": return .subscription
        case "culture", "hobby", "golf", "travel", "travel_lux", "art_buy", "broke_game", "working_hwe", "religion", "donation", "foundation": return .leisure
        case "kids", "kids_private", "edu_self", "pet": return .education
        case "wedding", "funeral", "baby", "gift_send", "working_parent", "alimony", "staff_bonus", "loan_give", "broke_lend": return .family
        case "fine_traffic", "fine_parking", "fine_misc", "tax_property1", "tax_property2", "tax_car", "tax_car2", "tax_resident", "tax_yearend_pay", "tax_compre", "tax_cpt", "pension", "nhi", "broke_cut", "broke_seize": return .tax
        case "accident_pay", "legal", "scam", "lost_phone", "realestate_buy": return .other
        default:
            if id.hasPrefix("ins_") { return .insurance }
            if id.hasPrefix("med_") { return .medical }
            if id.hasPrefix("loan_") || id.hasPrefix("fine_late") || id == "working_saving" { return .finance }
            return .other
        }
    }

    private static let keywords: [(SpendingCategory, [String])] = [
        (.cafe, ["스타벅스", "커피", "이디야", "투썸", "카페", "베이커리", "빽다방", "메가커피", "컴포즈"]),
        (.food, ["식당", "배달", "맥도날드", "버거", "치킨", "김밥", "도시락", "서브웨이", "맘스터치", "롯데리아", "본죽", "쿠팡이츠", "요기요", "우아한형제들", "식비", "점심", "저녁", "야식", "구내식당", "분식"]),
        (.mart, ["GS25", "CU", "세븐일레븐", "이마트", "홈플러스", "롯데마트", "컬리", "코스트코", "편의점", "마트", "장보기", "쿠팡(주)", "노브랜드"]),
        (.shopping, ["무신사", "올리브영", "유니클로", "자라", "다이소", "쇼핑", "지그재그", "29CM", "애플", "Apple", "디지털프라자", "하이마트", "샤넬", "루이비통", "에르메스", "롤렉스", "구찌", "디올", "까르띠에", "헤어", "네일", "미용"]),
        (.transport, ["택시", "카카오T", "우버", "타다", "티머니", "교통", "주유", "충전", "하이패스", "도로공사", "블루핸즈", "오토큐", "타이어", "따릉이", "지하철", "버스", "전용기", "VistaJet", "요트"]),
        (.housing, ["월세", "관리비", "관리사무소", "고시원", "전기", "가스", "수도", "인터넷", "SK텔레콤", "KT", "LG유플러스", "휴대폰", "통신", "세탁", "이사", "보증금", "집주인", "수리"]),
        (.subscription, ["넷플릭스", "디즈니", "유튜브", "티빙", "웨이브", "멜론", "스포티파이", "구독", "멤버십", "iCloud", "Google", "App Store", "Steam", "짐", "헬스"]),
        (.medical, ["병원", "의원", "약국", "치과", "내과", "안과", "피부과", "정형외과", "응급", "검진", "진료", "동물병원", "상담"]),
        (.insurance, ["보험", "화재", "생명", "라이프", "손해보험", "펫보험"]),
        (.leisure, ["CGV", "메가박스", "롯데시네마", "영화", "콘서트", "티켓", "교보문고", "예스24", "골프", "CC", "볼링", "클라이밍", "캠핑", "테니스", "여행", "항공", "호텔", "아고다", "야놀자", "리조트", "갤러리", "경매", "토토", "경마", "게임", "회식", "동호회", "헌금", "후원", "기부", "교회", "성당"]),
        (.education, ["학원", "유치원", "학습지", "태권도", "피아노", "어학", "강의", "인프런", "클래스101", "해커스", "국제학교", "유학", "과외", "승마", "키즈", "토이저러스", "펫", "사료"]),
        (.family, ["결혼", "축의금", "부의금", "조의금", "장례", "돌잔치", "출산", "용돈", "엄마", "아빠", "부모님", "장모", "장인", "시어머니", "시아버지", "생활비 보태", "양육비", "위자료", "빌려"]),
        (.tax, ["세금", "국세청", "홈택스", "위택스", "구청", "세무", "재산세", "자동차세", "주민세", "종합소득세", "과태료", "범칙금", "벌금", "국민연금", "건강보험", "압류", "집행관"]),
        (.finance, ["이자", "상환", "대출", "캐피탈", "카드론", "리볼빙", "연체", "적금", "예금", "청약", "카드 대금", "카드대금", "추심"]),
    ]

    /// Keyword fallback for transactions without a tag (user transfers, core flows, old data).
    static func classify(counterparty: String, memo: String, amount: Decimal, category: TransactionCategory) -> SpendingCategory {
        if amount > 0 { return .income }
        switch category {
        case .interest, .repayment, .maturity: return .finance
        case .adjustment: return .other
        default: break
        }
        let text = counterparty + " " + memo
        for (cat, words) in keywords where words.contains(where: { text.localizedCaseInsensitiveContains($0) }) { return cat }
        return category == .transfer ? .transfer : .other
    }
}

/// Pure monthly aggregation used by the 소비 리포트 screen and the widget.
struct SpendingReport: Equatable {
    struct Slice: Identifiable, Equatable {
        let category: SpendingCategory
        let amount: Decimal
        let count: Int
        var id: String { category.rawValue }
    }
    struct Merchant: Identifiable, Equatable {
        let name: String
        let amount: Decimal
        let count: Int
        var id: String { name }
    }

    let month: Date
    let spent: Decimal
    let income: Decimal
    let slices: [Slice]
    let merchants: [Merchant]
    let dailySpent: [Date: Decimal]
    let previousSpent: Decimal

    var changeRatio: Double? {
        guard previousSpent > 0 else { return nil }
        return ((spent - previousSpent) / previousSpent).doubleValue
    }

    static func category(of tx: Transaction) -> SpendingCategory {
        if let tag = tx.tag, let c = SpendingCategory(rawValue: tag) { return c }
        return SpendingCategory.classify(counterparty: tx.counterparty, memo: tx.memo, amount: tx.amount, category: tx.category)
    }

    static func build(transactions: [Transaction], month: Date, calendar: Calendar = .current) -> SpendingReport {
        func inMonth(_ d: Date, _ m: Date) -> Bool { calendar.isDate(d, equalTo: m, toGranularity: .month) }
        let rows = transactions.filter { inMonth($0.date, month) && $0.category != .adjustment && $0.tag != SpendingCategory.transfer.rawValue }
        var spent: Decimal = 0, income: Decimal = 0
        var byCat: [SpendingCategory: (Decimal, Int)] = [:]
        var byMerchant: [String: (Decimal, Int)] = [:]
        var daily: [Date: Decimal] = [:]
        for tx in rows {
            let cat = category(of: tx)
            if tx.amount > 0 {
                if cat == .income || tx.category != .transfer { income += tx.amount }
                continue
            }
            let a = -tx.amount
            spent += a
            byCat[cat, default: (0, 0)].0 += a; byCat[cat, default: (0, 0)].1 += 1
            byMerchant[tx.counterparty, default: (0, 0)].0 += a; byMerchant[tx.counterparty, default: (0, 0)].1 += 1
            daily[calendar.startOfDay(for: tx.date), default: 0] += a
        }
        let prevMonth = calendar.date(byAdding: .month, value: -1, to: month) ?? month
        let prev = transactions.filter { inMonth($0.date, prevMonth) && $0.amount < 0 && $0.category != .adjustment && $0.tag != SpendingCategory.transfer.rawValue }
            .reduce(Decimal(0)) { $0 - $1.amount }
        return SpendingReport(
            month: month, spent: spent, income: income,
            slices: byCat.map { Slice(category: $0.key, amount: $0.value.0, count: $0.value.1) }.sorted { $0.amount > $1.amount },
            merchants: byMerchant.map { Merchant(name: $0.key, amount: $0.value.0, count: $0.value.1) }.sorted { $0.amount > $1.amount },
            dailySpent: daily, previousSpent: prev)
    }
}

import Foundation

// MARK: - Money helpers

extension Decimal {
    /// Round to `scale` decimal places (bankers rounding).
    func rounded(scale: Int = 0) -> Decimal {
        var value = self
        var result = Decimal()
        NSDecimalRound(&result, &value, scale, .bankers)
        return result
    }

    var doubleValue: Double { NSDecimalNumber(decimal: self).doubleValue }

    /// Double → Decimal with float noise removed (8 dp keeps rate factors like 0.042/12 intact).
    init(double: Double) {
        guard double.isFinite else { self = 0; return }
        self = Decimal(double).rounded(scale: 8)
    }
}

enum Money {
    /// Mirrors Theme.wonSuffix of the active theme; BankStore keeps it in sync.
    nonisolated(unsafe) static var wonSuffix = false

    static func format(_ amount: Decimal, code: String, signed: Bool = false) -> String {
        if code == "KRW", wonSuffix {
            let f = NumberFormatter()
            f.numberStyle = .decimal
            f.maximumFractionDigits = 0
            let body = f.string(from: NSDecimalNumber(decimal: abs(amount))) ?? "\(amount)"
            let sign = amount < 0 ? "-" : (signed && amount > 0 ? "+" : "")
            return "\(sign)\(body)원"
        }
        let f = NumberFormatter()
        f.numberStyle = .currency
        f.currencyCode = code
        f.locale = code == "KRW" ? Locale(identifier: "ko_KR") : Locale(identifier: "en_US")
        f.maximumFractionDigits = code == "KRW" ? 0 : 2
        if signed { f.positivePrefix = "+" + f.positivePrefix }
        return f.string(from: NSDecimalNumber(decimal: amount)) ?? "\(amount)"
    }
}

// MARK: - Accounts

enum AccountKind: String, Codable, CaseIterable, Identifiable {
    case checking, installmentSavings, deposit, loan
    var id: String { rawValue }

    var title: String {
        switch self {
        case .checking: return "입출금"
        case .installmentSavings: return "적금"
        case .deposit: return "예금"
        case .loan: return "대출"
        }
    }

    var symbol: String {
        switch self {
        case .checking: return "creditcard.fill"
        case .installmentSavings: return "calendar.badge.plus"
        case .deposit: return "lock.shield.fill"
        case .loan: return "building.columns.fill"
        }
    }

    var isProduct: Bool { self != .checking }
}

/// Terms for interest-bearing products (savings, deposit, loan).
struct ProductTerms: Codable, Equatable {
    var annualRate: Double            // 0.035 == 3.5%
    var termMonths: Int
    var startDate: Date
    var principal: Decimal            // deposit: lump sum, loan: borrowed amount, savings: 0
    var monthlyAmount: Decimal        // savings: monthly deposit, loan: monthly payment, deposit: 0
    var linkedAccountID: UUID?        // checking account used for auto transfers
    var lastAccrualDate: Date         // interest accrued up to this date (daily)
    var lastInstallmentDate: Date     // last monthly transfer/repayment processed
    var accruedInterest: Decimal = 0  // interest earned but not yet paid out
    var isClosed: Bool = false        // matured / fully repaid

    var maturityDate: Date {
        Calendar.current.date(byAdding: .month, value: termMonths, to: startDate) ?? startDate
    }
}

struct Account: Identifiable, Codable, Equatable {
    var id = UUID()
    var name: String
    var kind: AccountKind
    var number: String
    var balance: Decimal
    var createdAt = Date()
    var terms: ProductTerms?
    var colorIndex: Int = 0

    static func generateNumber() -> String {
        func chunk(_ n: Int) -> String { (0..<n).map { _ in String(Int.random(in: 0...9)) }.joined() }
        return "\(chunk(3))-\(chunk(4))-\(chunk(6))"
    }
}

// MARK: - Transactions

enum TransactionCategory: String, Codable {
    case transfer, scheduled, interest, autoTransfer, repayment, adjustment, maturity
    var title: String {
        switch self {
        case .transfer: return "이체"
        case .scheduled: return "예약 입금"
        case .interest: return "이자"
        case .autoTransfer: return "자동이체"
        case .repayment: return "상환"
        case .adjustment: return "잔액 설정"
        case .maturity: return "만기"
        }
    }
}

struct Transaction: Identifiable, Codable, Equatable, Hashable {
    var id = UUID()
    var accountID: UUID
    var date: Date
    var amount: Decimal          // signed: + deposit, - withdrawal (for loan accounts + means debt increased)
    var counterparty: String
    var memo: String
    var category: TransactionCategory
    var balanceAfter: Decimal
    var tag: String? = nil          // SpendingCategory raw value when the generator knows it
}

// MARK: - Scheduled events ("Mom sends ₩5,000,000 at 4pm")

enum RepeatRule: String, Codable, CaseIterable, Identifiable {
    case none, daily, weekly, monthly
    var id: String { rawValue }
    var title: String {
        switch self {
        case .none: return "한 번만"
        case .daily: return "매일"
        case .weekly: return "매주"
        case .monthly: return "매월"
        }
    }

    func next(after date: Date) -> Date? {
        let cal = Calendar.current
        switch self {
        case .none: return nil
        case .daily: return cal.date(byAdding: .day, value: 1, to: date)
        case .weekly: return cal.date(byAdding: .weekOfYear, value: 1, to: date)
        case .monthly: return cal.date(byAdding: .month, value: 1, to: date)
        }
    }
}

struct ScheduledEvent: Identifiable, Codable, Equatable {
    var id = UUID()
    var accountID: UUID
    var counterparty: String      // "엄마"
    var amount: Decimal           // positive = incoming, negative = outgoing
    var memo: String              // "용돈"
    var nextFireDate: Date
    var repeatRule: RepeatRule
    var isEnabled = true
    var lastFiredAt: Date?
    var toAccountID: UUID? = nil  // set = internal transfer: amount leaves accountID and lands here

    var isIncoming: Bool { amount >= 0 }
    var isInternal: Bool { toAccountID != nil }
}

// MARK: - Settings / root document

enum Appearance: String, Codable, CaseIterable, Identifiable {
    case system, light, dark
    var id: String { rawValue }
    var title: String {
        switch self {
        case .system: return "시스템"
        case .light: return "라이트"
        case .dark: return "다크"
        }
    }
}

struct BankSettings: Codable, Equatable {
    var themeID: String = "nova"
    var appearance: Appearance = .system
    var currencyCode: String = "KRW"
    var ownerName: String = "홍길동"
    var interestTaxRate: Double = 0.154     // 이자소득세 15.4% (0이면 비과세)
    var hideBalances: Bool = false
    var hasOnboarded: Bool = false
    var appLock: Bool = false
    // User-editable institution names. Empty = theme default.
    var bankName: String = ""
    var brokerName: String = ""
    var cardName: String = ""
    var capitalName: String = ""
    var ageVerified: Bool = false          // 만 18세 확인 (블랙잭 테이블)
    var proBonusGranted: Bool = false      // one-time Pro 보너스 게임 머니 credited

    var brand: Brand { Brand(settings: self) }

    init() {}

    // Tolerant decoding so adding a field never wipes saved data.
    init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        themeID = try c.decodeIfPresent(String.self, forKey: .themeID) ?? themeID
        appearance = try c.decodeIfPresent(Appearance.self, forKey: .appearance) ?? appearance
        currencyCode = try c.decodeIfPresent(String.self, forKey: .currencyCode) ?? currencyCode
        ownerName = try c.decodeIfPresent(String.self, forKey: .ownerName) ?? ownerName
        interestTaxRate = try c.decodeIfPresent(Double.self, forKey: .interestTaxRate) ?? interestTaxRate
        hideBalances = try c.decodeIfPresent(Bool.self, forKey: .hideBalances) ?? hideBalances
        hasOnboarded = try c.decodeIfPresent(Bool.self, forKey: .hasOnboarded) ?? hasOnboarded
        appLock = try c.decodeIfPresent(Bool.self, forKey: .appLock) ?? appLock
        bankName = try c.decodeIfPresent(String.self, forKey: .bankName) ?? bankName
        brokerName = try c.decodeIfPresent(String.self, forKey: .brokerName) ?? brokerName
        cardName = try c.decodeIfPresent(String.self, forKey: .cardName) ?? cardName
        capitalName = try c.decodeIfPresent(String.self, forKey: .capitalName) ?? capitalName
        ageVerified = try c.decodeIfPresent(Bool.self, forKey: .ageVerified) ?? ageVerified
        proBonusGranted = try c.decodeIfPresent(Bool.self, forKey: .proBonusGranted) ?? proBonusGranted
    }
}

/// Resolved institution names. Templates use {bank} {broker} {card} {capital} placeholders.
struct Brand: Equatable {
    let bank: String
    let broker: String
    let card: String
    let capital: String

    static let defaultCard = AppVariant.isStore ? "미니멀카드" : "신한카드"
    static let defaultCapital = AppVariant.isStore ? "미니멀캐피탈" : "OK캐피탈"

    init(settings: BankSettings) {
        let theme = Theme.named(settings.themeID)
        // App Store build: a typed real bank name (e.g. 신한은행) is shown as its fictional counterpart
        bank = StoreNames.fix(settings.bankName.isEmpty ? theme.bankName : settings.bankName)
        broker = StoreNames.fix(settings.brokerName.isEmpty ? theme.brokerName : settings.brokerName)
        card = StoreNames.fix(settings.cardName.isEmpty ? Brand.defaultCard : settings.cardName)
        capital = StoreNames.fix(settings.capitalName.isEmpty ? Brand.defaultCapital : settings.capitalName)
    }

    func resolve(_ text: String) -> String {
        guard text.contains("{") else { return text }
        return text.replacingOccurrences(of: "{bank}", with: bank)
            .replacingOccurrences(of: "{broker}", with: broker)
            .replacingOccurrences(of: "{card}", with: card)
            .replacingOccurrences(of: "{capital}", with: capital)
    }
}

struct BankData: Codable, Equatable {
    var accounts: [Account] = []
    var transactions: [Transaction] = []
    var events: [ScheduledEvent] = []
    var settings = BankSettings()
    var portfolio = Portfolio.starter()
    var lifestyle = LifestyleState()
    var clock = SimClock()
    var challenge = ChallengeState()
    var goals: [SavingsGoal] = []
    var achievements: [String: Date] = [:]
    var game = GameState()
    var jobs = JobsState()             // 알바 mini games: plays, bests, today's shifts
    var life = LifeEventState()        // 교통사고·단속·구속·실직 and the current 부상/수감 status

    init() {}

    // Tolerant decoding: a new top-level field must never throw away the user's saved bank.
    init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        accounts = try c.decodeIfPresent([Account].self, forKey: .accounts) ?? []
        transactions = try c.decodeIfPresent([Transaction].self, forKey: .transactions) ?? []
        events = try c.decodeIfPresent([ScheduledEvent].self, forKey: .events) ?? []
        settings = try c.decodeIfPresent(BankSettings.self, forKey: .settings) ?? BankSettings()
        portfolio = try c.decodeIfPresent(Portfolio.self, forKey: .portfolio) ?? Portfolio.starter()
        lifestyle = try c.decodeIfPresent(LifestyleState.self, forKey: .lifestyle) ?? LifestyleState()
        clock = try c.decodeIfPresent(SimClock.self, forKey: .clock) ?? SimClock()
        challenge = try c.decodeIfPresent(ChallengeState.self, forKey: .challenge) ?? ChallengeState()
        goals = try c.decodeIfPresent([SavingsGoal].self, forKey: .goals) ?? []
        achievements = try c.decodeIfPresent([String: Date].self, forKey: .achievements) ?? [:]
        game = try c.decodeIfPresent(GameState.self, forKey: .game) ?? GameState()
        jobs = try c.decodeIfPresent(JobsState.self, forKey: .jobs) ?? JobsState()
        life = try c.decodeIfPresent(LifeEventState.self, forKey: .life) ?? LifeEventState()
    }

    static func starter() -> BankData {
        var d = BankData()
        var main = Account(name: "입출금통장", kind: .checking, number: Account.generateNumber(), balance: 1_250_000)
        main.colorIndex = 0
        d.accounts = [main]
        d.transactions = [Transaction(accountID: main.id, date: Date(), amount: 1_250_000,
                                      counterparty: d.settings.brand.bank, memo: "시작 잔액", category: .adjustment,
                                      balanceAfter: main.balance)]
        return d
    }
}

import Foundation

// MARK: - Quest content types (filled by QuestLibrary, generated)

enum QuestCategory: String, Codable, CaseIterable {
    case career, insurance, housing, invest, business, family, education, health, luxury, temptation, promotion, tax, finance
    var title: String {
        switch self {
        case .career: return "커리어"
        case .insurance: return "보험"
        case .housing: return "주거"
        case .invest: return "투자"
        case .business: return "사업"
        case .family: return "가족"
        case .education: return "교육"
        case .health: return "건강"
        case .luxury: return "소비"
        case .temptation: return "유혹"
        case .promotion: return "승급"
        case .tax: return "세금"
        case .finance: return "금융"
        }
    }
    var symbol: String {
        switch self {
        case .career: return "briefcase.fill"
        case .insurance: return "shield.fill"
        case .housing: return "house.fill"
        case .invest: return "chart.line.uptrend.xyaxis"
        case .business: return "storefront.fill"
        case .family: return "person.2.fill"
        case .education: return "graduationcap.fill"
        case .health: return "heart.fill"
        case .luxury: return "bag.fill"
        case .temptation: return "exclamationmark.triangle.fill"
        case .promotion: return "arrow.up.circle.fill"
        case .tax: return "doc.text.fill"
        case .finance: return "building.columns.fill"
        }
    }
}

/// A recurring flow added by a choice. `months == nil` means forever (until cancelled by another choice).
struct QuestFlow {
    let template: FlowTemplate
    var months: Int? = nil
}

/// Chance-based outcome resolved months after the choice (delivered as a scheduled event).
struct Gamble {
    let chance: Double
    let win: Decimal            // cash delivered on success (0 = none)
    let lose: Decimal           // cash delivered on failure (negative or 0)
    let delayMonths: Int
    let winText: String
    let loseText: String
    let counterparty: String
    var winFlow: QuestFlow? = nil
    var loseFlow: QuestFlow? = nil
    var loseSalaryPause: Int = 0
    var loseFine: Decimal = 0
}

/// Starting a business: after `delayMonths` it either settles into winFlow or bleeds via loseFlow.
struct Venture {
    let chance: Double
    let winFlow: QuestFlow
    let loseFlow: QuestFlow?
    let winText: String
    let loseText: String
    let delayMonths: Int
}

struct QuestEffect {
    var cash: Decimal = 0
    var flow: QuestFlow? = nil
    var flow2: QuestFlow? = nil
    var cancelFlowMemo: String? = nil        // removes active quest flows whose memo contains this
    var salaryMultiplier: Double = 1
    var counterSalary: Double? = nil         // gamble win → salary × this (역제안)
    var successSalary: Double? = nil         // gamble win → salary × this (승진/교육)
    var employer: String? = nil
    var jobTitle: String? = nil
    var expenseMultiplier: Double = 1
    var medicalMultiplier: Double = 1
    var gamble: Gamble? = nil
    var venture: Venture? = nil
    var promoteTo: WealthTier? = nil
    var netWorthMultiplier: Double? = nil    // immediate hit on checking balance (이혼 재산분할 등)
    var loseNetWorthMultiplier: Double? = nil
    var loseDemoteTo: WealthTier? = nil
    var salaryPauseMonths: Int = 0
    var loseSalaryPause: Int = 0
    var resist: Decimal = 0                  // "참는다" → logged into 무지출 챌린지
    var afterText: String? = nil
}

struct QuestOption {
    let title: String
    let note: String
    let effect: QuestEffect
}

struct Quest: Identifiable {
    let id: String
    let category: QuestCategory
    let tiers: [WealthTier]
    let minAge: Int
    let maxAge: Int
    let once: Bool
    let promotion: WealthTier?
    let requiredNetWorth: Decimal?
    let title: String
    let detail: String
    let options: [QuestOption]
}

extension QuestLibrary {
    static func quest(_ id: String) -> Quest? { byID[id] }
    static let byID: [String: Quest] = Dictionary(uniqueKeysWithValues: all.map { ($0.id, $0) })
}

// MARK: - Game state (persisted)

/// Recurring flow created by a choice, with optional expiry.
struct ActiveFlow: Codable, Equatable, Identifiable {
    var id = UUID()
    var questID: String
    var counterparty: String
    var memo: String
    var amount: Decimal
    var day: Int
    var category: TransactionCategory
    var hour: Int
    var until: Date? = nil
}

struct QuestAnswer: Codable, Equatable, Identifiable {
    var id = UUID()
    var questID: String
    var optionIndex: Int
    var date: Date
    var title: String
    var choice: String
    var outcome: String? = nil        // resolved text (gamble/venture) or immediate note
    var outcomeDate: Date? = nil
    var cash: Decimal = 0             // immediate cash effect
}

/// Pending delayed consequence of a choice.
struct PendingOutcome: Codable, Equatable, Identifiable {
    var id = UUID()
    var answerID: UUID
    var questID: String
    var fireDate: Date
    var success: Bool
    var cash: Decimal
    var text: String
    var counterparty: String
    var flow: ActiveFlow? = nil
    var flowMonths: Int? = nil
    var salaryPause: Int = 0
    var salaryMultiplier: Double = 1
    var fine: Decimal = 0
    var netWorthMultiplier: Double? = nil
    var demoteTo: WealthTier? = nil
}

struct GameState: Codable, Equatable {
    var birthDate: Date? = nil
    var unlockedTiers: [WealthTier] = [.broke, .working]
    var answers: [QuestAnswer] = []
    var pendingQuestID: String? = nil
    var pendingSince: Date? = nil
    var pendingExpires: Date? = nil
    var nextQuestDate: Date? = nil
    var nextQuestID: String? = nil          // chosen ahead so the notification can name it
    var outcomes: [PendingOutcome] = []
    var flows: [ActiveFlow] = []
    var salaryMultiplier: Double = 1
    var expenseMultiplier: Double = 1
    var medicalMultiplier: Double = 1
    var employer: String? = nil
    var jobTitle: String? = nil
    var salaryPausedUntil: Date? = nil
    var endedAt: Date? = nil
    var lifespanYears: Int = 80
    var startAge: Int = 22

    static let questIntervalDays = 5...14
    static let questWindowDays = 14

    func age(at date: Date, calendar: Calendar = .current) -> Int {
        guard let b = birthDate else { return startAge }
        return max(0, calendar.dateComponents([.year], from: b, to: date).year ?? 0)
    }
    var isEnded: Bool { endedAt != nil }
    func answered(_ questID: String) -> Bool { answers.contains { $0.questID == questID } }
    func isUnlocked(_ tier: WealthTier) -> Bool { unlockedTiers.contains(tier) }
    var activeFlows: [ActiveFlow] { flows }

    static func nextTier(after t: WealthTier) -> WealthTier? {
        let all = WealthTier.allCases
        guard let i = all.firstIndex(of: t), i + 1 < all.count else { return nil }
        return all[i + 1]
    }
    /// Net worth needed before the promotion quest to `tier` can appear.
    static func threshold(for tier: WealthTier) -> Decimal {
        switch tier {
        case .broke: return 0
        case .working: return 0
        case .middle: return 50_000_000
        case .rich: return 1_000_000_000
        case .tycoon: return 50_000_000_000
        }
    }
}

import Foundation

/// How often an automatic lifestyle flow happens.
enum FlowFrequency: Codable, Equatable {
    case monthly(day: Int)          // on this day of month
    case weekly(weekday: Int)       // 1 = Sunday ... 7 = Saturday
    case daily(chance: Double)      // probability per day
}

/// One recurring cash flow that a wealth tier generates (e.g. "배당금", "캐피탈 상환").
struct FlowTemplate: Codable, Equatable {
    var counterparty: String
    var memo: String
    var minAmount: Decimal
    var maxAmount: Decimal
    var frequency: FlowFrequency
    var category: TransactionCategory = .transfer
    var hour: Int = 10
    var isSalary: Bool = false     // grows with years of service; everything else follows inflation
}

/// Wealth tiers. Picking one drives the stream of automatic transactions.
enum WealthTier: String, Codable, CaseIterable, Identifiable {
    case broke, working, middle, rich, tycoon
    var id: String { rawValue }

    var title: String {
        switch self {
        case .broke: return "거지"
        case .working: return "서민"
        case .middle: return "중산층"
        case .rich: return "부자"
        case .tycoon: return "재벌"
        }
    }
    var emoji: String {
        switch self {
        case .broke: return "🥲"
        case .working: return "🧑‍🔧"
        case .middle: return "🏠"
        case .rich: return "🛥️"
        case .tycoon: return "🏦"
        }
    }
    var summary: String {
        switch self {
        case .broke: return "알바비가 들어오자마자 캐피탈 상환·연체이자·소액결제로 전부 빠져나갑니다. 잔액은 늘 마이너스 근처."
        case .working: return "월급 280만 원, 월세·카드값·대출이자 나가면 남는 게 별로 없습니다."
        case .middle: return "월급 620만 원, 아파트 관리비·학원비·적금. 조금씩 쌓입니다."
        case .rich: return "급여 2,500만 원에 임대료·배당금이 매주 들어옵니다. 지출은 신경 안 씁니다."
        case .tycoon: return "법인 급여 1.2억, 투자사 수익배분·배당·채권 이자·임대료가 매일 들어옵니다. 나가는 돈은 세금뿐."
        }
    }
    /// Yearly raise applied to salary-type flows on each anniversary.
    var annualRaise: Double {
        switch self {
        case .broke: return 0.012      // 최저임금 수준 인상
        case .working: return 0.035
        case .middle: return 0.048
        case .rich: return 0.07
        case .tycoon: return 0.09
        }
    }
    /// Promotion every N years multiplies salary by `promotionBump`.
    var promotionEveryYears: Int {
        switch self {
        case .broke: return 100
        case .working: return 6
        case .middle: return 4
        case .rich: return 3
        case .tycoon: return 2
        }
    }
    var promotionBump: Double {
        switch self {
        case .broke: return 1.0
        case .working: return 1.12
        case .middle: return 1.18
        case .rich: return 1.25
        case .tycoon: return 1.35
        }
    }
    static let inflation = 0.025     // yearly growth of expenses

    /// Suggested starting checking balance when applying the tier.
    var startingBalance: Decimal {
        switch self {
        case .broke: return 13_240
        case .working: return 412_000
        case .middle: return 8_950_000
        case .rich: return 187_400_000
        case .tycoon: return 4_830_000_000
        }
    }

    var flows: [FlowTemplate] {
        switch self {
        case .broke: return [
            FlowTemplate(counterparty: "CU 신림역점", memo: "알바 주급", minAmount: 180_000, maxAmount: 260_000, frequency: .weekly(weekday: 6), hour: 18, isSalary: true),
            FlowTemplate(counterparty: "{capital}", memo: "대출 원금 상환", minAmount: -210_000, maxAmount: -210_000, frequency: .monthly(day: 5), category: .repayment, hour: 9),
            FlowTemplate(counterparty: "(주)러시앤캐시", memo: "대출 이자", minAmount: -94_000, maxAmount: -94_000, frequency: .monthly(day: 12), category: .interest, hour: 9),
            FlowTemplate(counterparty: "{capital}", memo: "연체 이자", minAmount: -31_000, maxAmount: -47_000, frequency: .weekly(weekday: 2), category: .interest, hour: 8),
            FlowTemplate(counterparty: "KT M모바일 소액결제", memo: "자동이체", minAmount: -89_000, maxAmount: -89_000, frequency: .monthly(day: 25), hour: 7),
            FlowTemplate(counterparty: "하늘고시원 신림점", memo: "월세", minAmount: -280_000, maxAmount: -280_000, frequency: .monthly(day: 1), hour: 10),
            FlowTemplate(counterparty: "GS25 신림점", memo: "컵라면·삼각김밥", minAmount: -3_200, maxAmount: -8_900, frequency: .daily(chance: 0.9), hour: 21),
            FlowTemplate(counterparty: "친구", memo: "빌린 돈 갚음", minAmount: -20_000, maxAmount: -50_000, frequency: .daily(chance: 0.1), hour: 22),
        ]
        case .working: return [
            FlowTemplate(counterparty: "(주)대성물류", memo: "급여", minAmount: 2_800_000, maxAmount: 2_800_000, frequency: .monthly(day: 25), hour: 9, isSalary: true),
            FlowTemplate(counterparty: "집주인", memo: "월세", minAmount: -650_000, maxAmount: -650_000, frequency: .monthly(day: 1), hour: 10),
            FlowTemplate(counterparty: "{card}", memo: "카드 대금", minAmount: -980_000, maxAmount: -1_340_000, frequency: .monthly(day: 14), hour: 8),
            FlowTemplate(counterparty: "햇살론", memo: "대출 이자", minAmount: -112_000, maxAmount: -112_000, frequency: .monthly(day: 10), category: .interest, hour: 9),
            FlowTemplate(counterparty: "SK텔레콤", memo: "휴대폰 요금", minAmount: -69_000, maxAmount: -69_000, frequency: .monthly(day: 20), hour: 7),
            FlowTemplate(counterparty: "한솥도시락 가산디지털점", memo: "식비", minAmount: -7_000, maxAmount: -18_000, frequency: .daily(chance: 0.8), hour: 12),
            FlowTemplate(counterparty: "티머니", memo: "교통카드 충전", minAmount: -50_000, maxAmount: -50_000, frequency: .weekly(weekday: 1), hour: 8),
        ]
        case .middle: return [
            FlowTemplate(counterparty: "(주)한빛전자", memo: "급여", minAmount: 6_200_000, maxAmount: 6_200_000, frequency: .monthly(day: 25), hour: 9, isSalary: true),
            FlowTemplate(counterparty: "래미안 관리사무소", memo: "관리비", minAmount: -310_000, maxAmount: -420_000, frequency: .monthly(day: 5), hour: 10),
            FlowTemplate(counterparty: "{card}", memo: "카드 대금", minAmount: -1_900_000, maxAmount: -2_600_000, frequency: .monthly(day: 15), hour: 8),
            FlowTemplate(counterparty: "{bank} 주택담보대출", memo: "원리금 상환", minAmount: -1_420_000, maxAmount: -1_420_000, frequency: .monthly(day: 10), category: .repayment, hour: 9),
            FlowTemplate(counterparty: "CMS 수학학원 목동점", memo: "학원비", minAmount: -480_000, maxAmount: -480_000, frequency: .monthly(day: 3), hour: 11),
            FlowTemplate(counterparty: "{bank} 자유적금", memo: "자동이체", minAmount: -500_000, maxAmount: -500_000, frequency: .monthly(day: 26), category: .autoTransfer, hour: 9),
            FlowTemplate(counterparty: "이마트 목동점", memo: "장보기", minAmount: -90_000, maxAmount: -210_000, frequency: .weekly(weekday: 7), hour: 15),
            FlowTemplate(counterparty: "스타벅스 역삼대로점", memo: "커피", minAmount: -5_900, maxAmount: -12_000, frequency: .daily(chance: 0.6), hour: 9),
        ]
        case .rich: return [
            FlowTemplate(counterparty: "(주)케이원홀딩스", memo: "급여", minAmount: 25_000_000, maxAmount: 25_000_000, frequency: .monthly(day: 25), hour: 9, isSalary: true),
            FlowTemplate(counterparty: "강남 오피스 임차인", memo: "임대료", minAmount: 8_500_000, maxAmount: 8_500_000, frequency: .monthly(day: 1), hour: 10),
            FlowTemplate(counterparty: "한남동 상가", memo: "임대료", minAmount: 4_200_000, maxAmount: 4_200_000, frequency: .monthly(day: 5), hour: 10),
            FlowTemplate(counterparty: "{broker}", memo: "배당금 입금", minAmount: 1_200_000, maxAmount: 6_800_000, frequency: .weekly(weekday: 3), category: .interest, hour: 15),
            FlowTemplate(counterparty: "{bank}", memo: "예금 이자", minAmount: 640_000, maxAmount: 640_000, frequency: .monthly(day: 20), category: .interest, hour: 9),
            FlowTemplate(counterparty: "현대카드 Black", memo: "카드 대금", minAmount: -9_000_000, maxAmount: -17_000_000, frequency: .monthly(day: 14), hour: 8),
            FlowTemplate(counterparty: "국세청", memo: "종합소득세 중간예납", minAmount: -12_000_000, maxAmount: -12_000_000, frequency: .monthly(day: 30), hour: 10),
            FlowTemplate(counterparty: "남서울CC 클럽하우스", memo: "결제", minAmount: -300_000, maxAmount: -1_800_000, frequency: .daily(chance: 0.5), hour: 20),
        ]
        case .tycoon: return [
            FlowTemplate(counterparty: "(주)VX그룹", memo: "법인 급여", minAmount: 120_000_000, maxAmount: 120_000_000, frequency: .monthly(day: 25), hour: 9, isSalary: true),
            FlowTemplate(counterparty: "VX벤처스", memo: "투자 수익 배분", minAmount: 35_000_000, maxAmount: 280_000_000, frequency: .weekly(weekday: 2), hour: 11),
            FlowTemplate(counterparty: "{broker}", memo: "배당금 입금", minAmount: 12_000_000, maxAmount: 95_000_000, frequency: .daily(chance: 0.55), category: .interest, hour: 15),
            FlowTemplate(counterparty: "{bank} PB", memo: "채권 이자", minAmount: 4_800_000, maxAmount: 21_000_000, frequency: .daily(chance: 0.7), category: .interest, hour: 9),
            FlowTemplate(counterparty: "테헤란로 타워", memo: "임대료", minAmount: 310_000_000, maxAmount: 310_000_000, frequency: .monthly(day: 1), hour: 10),
            FlowTemplate(counterparty: "제주 리조트", memo: "운영 수익", minAmount: 48_000_000, maxAmount: 130_000_000, frequency: .weekly(weekday: 5), hour: 17),
            FlowTemplate(counterparty: "국세청", memo: "종합소득세", minAmount: -380_000_000, maxAmount: -380_000_000, frequency: .monthly(day: 30), hour: 10),
            FlowTemplate(counterparty: "VX재단", memo: "기부", minAmount: -50_000_000, maxAmount: -200_000_000, frequency: .monthly(day: 15), hour: 14),
            FlowTemplate(counterparty: "VistaJet", memo: "운항 비용", minAmount: -28_000_000, maxAmount: -64_000_000, frequency: .weekly(weekday: 6), hour: 19),
        ]
        }
    }
}

struct LifestyleState: Codable, Equatable {
    var tier: WealthTier? = nil
    var accountID: UUID? = nil          // checking account the flows hit
    var debtAccountID: UUID? = nil      // 거지: capital-loan account that keeps growing
    var tierStart: Date = Date()        // anniversary base for raises / inflation
    var lastProcessed: Date = Date()
    var seed: UInt64 = UInt64.random(in: 1...UInt64.max)
}

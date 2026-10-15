import Foundation

/// Deterministic RNG so replaying the same days yields the same transactions.
struct SplitMix64: RandomNumberGenerator {
    private var state: UInt64
    init(seed: UInt64) { state = seed }
    mutating func next() -> UInt64 {
        state &+= 0x9E3779B97F4A7C15
        var z = state
        z = (z ^ (z >> 30)) &* 0xBF58476D1CE4E5B9
        z = (z ^ (z >> 27)) &* 0x94D049BB133111EB
        return z ^ (z >> 31)
    }
}

/// Generates the life of the selected wealth tier, day by day, from the last processed
/// instant up to `now`:
///  - core flows (salary, rent, loans...) — salary grows with years of service and promotions,
///    everything else drifts with inflation
///  - the EventLibrary (insurance, utilities, medical, accidents, gifts, taxes, bonuses...)
///  - 거지 only: a capital loan that auto-tops-up the account and compounds with late interest,
///    so the ending is always a mountain of debt.
enum LifestyleEngine {
    static let maxDays = 120
    static let debtMonthlyRate = 0.031        // ~연 37% (이자 + 연체 가산)
    static let debtLateFee: Decimal = 50_000

    static func process(now: Date = Date(), data: inout BankData) {
        guard let tier = data.lifestyle.tier,
              let accountID = data.lifestyle.accountID,
              data.accounts.contains(where: { $0.id == accountID }) else { return }
        let cal = Calendar.current
        var day = cal.startOfDay(for: data.lifestyle.lastProcessed)
        let endDay = cal.startOfDay(for: now)
        var processed = data.lifestyle.lastProcessed
        var guardCount = 0
        let specs = EventLibrary.specs(for: tier)
        let scale = EventLibrary.tierScale[tier] ?? 1

        while day <= endDay, guardCount < maxDays {
            guardCount += 1
            let comps = cal.dateComponents([.year, .month, .day, .weekday], from: day)
            let dayKey = UInt64(comps.year! * 10_000 + comps.month! * 100 + comps.day!)
            var rng = SplitMix64(seed: data.lifestyle.seed ^ (dayKey &* 0x2545F4914F6CDD1D))
            let lastDay = cal.range(of: .day, in: .month, for: day)?.count ?? 28
            let years = max(0, cal.dateComponents([.year], from: data.lifestyle.tierStart, to: day).year ?? 0)
            let inflation = pow(1 + WealthTier.inflation, Double(years))
            let salaryFactor = pow(1 + tier.annualRaise, Double(years)) * pow(tier.promotionBump, Double(years / tier.promotionEveryYears))

            func emit(_ amount: Decimal, hour: Int, second: Int, counterparty: String, memo: String, category: TransactionCategory) {
                let minute = Int.random(in: 0..<60, using: &rng)
                guard let at = cal.date(bySettingHour: min(23, max(0, hour)), minute: minute, second: second % 60, of: day),
                      at > data.lifestyle.lastProcessed, at <= now else { return }
                Ledger.apply(amount: amount, to: accountID, date: at, counterparty: counterparty, memo: memo, category: category, data: &data)
                processed = max(processed, at)
            }

            // 1. core flows
            for (i, flow) in tier.flows.enumerated() {
                let fires: Bool
                switch flow.frequency {
                case .monthly(let d): fires = comps.day == min(d, lastDay)
                case .weekly(let wd): fires = comps.weekday == wd
                case .daily(let chance): fires = Double.random(in: 0..<1, using: &rng) < chance
                }
                guard fires else { continue }
                let factor = flow.isSalary ? salaryFactor : inflation
                let amount = randomAmount(flow.minAmount, flow.maxAmount, factor: factor, using: &rng)
                let memo = flow.isSalary && years > 0 ? "\(flow.memo) (\(years + 1)년차)" : flow.memo
                emit(amount, hour: flow.hour, second: i, counterparty: flow.counterparty, memo: memo, category: flow.category)
            }

            // 2. life events library
            for (i, spec) in specs.enumerated() {
                let fires: Bool
                switch spec.kind {
                case .daily(let chance): fires = Double.random(in: 0..<1, using: &rng) < chance
                case .monthly(let d): fires = comps.day == min(d, lastDay)
                case .yearly(let m, let d): fires = comps.month == m && comps.day == min(d, lastDay)
                }
                guard fires else { continue }
                let cp = spec.counterparties[Int.random(in: 0..<spec.counterparties.count, using: &rng)]
                let memo = spec.memos[Int.random(in: 0..<spec.memos.count, using: &rng)]
                let amount = randomAmount(Decimal(spec.baseMin), Decimal(spec.baseMax), factor: scale * inflation, using: &rng)
                emit(amount, hour: spec.hour, second: 10 + i, counterparty: cp, memo: memo, category: spec.category)
            }

            // 3. 거지 debt spiral (end of day)
            if tier == .broke, let debtID = data.lifestyle.debtAccountID,
               let di = data.accounts.firstIndex(where: { $0.id == debtID }),
               let ci = data.accounts.firstIndex(where: { $0.id == accountID }),
               let eod = cal.date(bySettingHour: 23, minute: 59, second: 0, of: day),
               eod > data.lifestyle.lastProcessed, eod <= now {
                let debtName = data.accounts[di].name
                if comps.day == 1 {
                    let debt = data.accounts[di].balance
                    if debt > 0 {
                        let interest = (debt * Decimal(double: debtMonthlyRate)).rounded()
                        Ledger.apply(amount: interest, to: debtID, date: eod, counterparty: debtName,
                                     memo: "대출 이자 + 연체 가산이자", category: .interest, data: &data)
                        let due = min((debt * Decimal(double: 0.04)).rounded(), debt + interest)
                        if data.accounts[ci].balance >= due, due > 0 {
                            Ledger.apply(amount: -due, to: accountID, date: eod, counterparty: debtName, memo: "최소 상환액 자동출금", category: .repayment, data: &data)
                            Ledger.apply(amount: -due, to: debtID, date: eod, counterparty: "본인 계좌", memo: "최소 상환", category: .repayment, data: &data)
                        } else {
                            Ledger.apply(amount: debtLateFee, to: debtID, date: eod, counterparty: "채권관리팀",
                                         memo: "연체 등록 · 연체 수수료", category: .interest, data: &data)
                            Ledger.apply(amount: 0, to: accountID, date: eod, counterparty: debtName,
                                         memo: "상환 실패 (잔액 부족) — 연체 \(Money.format(due, code: data.settings.currencyCode))", category: .repayment, data: &data)
                        }
                    }
                }
                let bal = data.accounts[ci].balance
                if bal < 0 {
                    let topUp = ((-bal + 30_000) / 10_000).rounded() * 10_000
                    Ledger.apply(amount: topUp, to: accountID, date: eod, counterparty: debtName, memo: "추가 대출 실행 (잔액 부족 자동)", category: .transfer, data: &data)
                    Ledger.apply(amount: topUp, to: debtID, date: eod, counterparty: "본인 계좌", memo: "추가 대출 (원금 증가)", category: .transfer, data: &data)
                }
                processed = max(processed, eod)
            }

            guard let next = cal.date(byAdding: .day, value: 1, to: day) else { break }
            day = next
            if day <= endDay, guardCount >= maxDays {
                data.lifestyle.lastProcessed = max(processed, day.addingTimeInterval(-1))
                return
            }
        }
        data.lifestyle.lastProcessed = max(processed, now)
    }

    private static func randomAmount(_ lo: Decimal, _ hi: Decimal, factor: Double, using rng: inout SplitMix64) -> Decimal {
        let a = min(lo.doubleValue, hi.doubleValue), b = max(lo.doubleValue, hi.doubleValue)
        let v = (a == b ? a : Double.random(in: a...b, using: &rng)) * factor
        let unit: Double = abs(v) >= 100_000_000 ? 1_000_000 : abs(v) >= 1_000_000 ? 10_000 : abs(v) >= 10_000 ? 1_000 : 100
        return Decimal(double: (v / unit).rounded() * unit)
    }
}

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

/// Generates the automatic cash flows of the selected wealth tier, day by day,
/// from the last processed date up to now (also used to back-fill history).
enum LifestyleEngine {
    static let maxDays = 120

    static func process(now: Date = Date(), data: inout BankData) {
        guard let tier = data.lifestyle.tier,
              let accountID = data.lifestyle.accountID,
              data.accounts.contains(where: { $0.id == accountID }) else { return }
        let cal = Calendar.current
        var day = cal.startOfDay(for: data.lifestyle.lastProcessed)
        let endDay = cal.startOfDay(for: now)
        var processed = data.lifestyle.lastProcessed
        var guardCount = 0

        while day <= endDay, guardCount < maxDays {
            guardCount += 1
            let comps = cal.dateComponents([.year, .month, .day, .weekday], from: day)
            let dayKey = UInt64(comps.year! * 10_000 + comps.month! * 100 + comps.day!)
            var rng = SplitMix64(seed: data.lifestyle.seed ^ (dayKey &* 0x2545F4914F6CDD1D))

            for (i, flow) in tier.flows.enumerated() {
                let fires: Bool
                switch flow.frequency {
                case .monthly(let d):
                    let lastDay = cal.range(of: .day, in: .month, for: day)?.count ?? 28
                    fires = comps.day == min(d, lastDay)
                case .weekly(let wd): fires = comps.weekday == wd
                case .daily(let chance): fires = Double.random(in: 0..<1, using: &rng) < chance
                }
                guard fires else { continue }
                let minute = Int.random(in: 0..<60, using: &rng)
                guard let at = cal.date(bySettingHour: flow.hour, minute: minute, second: i, of: day) else { continue }
                // only new moments: after the last processed instant and not in the future
                guard at > data.lifestyle.lastProcessed, at <= now else { continue }
                let amount = randomAmount(flow.minAmount, flow.maxAmount, using: &rng)
                Ledger.apply(amount: amount, to: accountID, date: at, counterparty: flow.counterparty,
                             memo: flow.memo, category: flow.category, data: &data)
                processed = max(processed, at)
            }
            guard let next = cal.date(byAdding: .day, value: 1, to: day) else { break }
            day = next
            if day <= endDay, guardCount >= maxDays {
                // budget exhausted: resume from the start of the next day on the next tick
                data.lifestyle.lastProcessed = max(processed, day.addingTimeInterval(-1))
                return
            }
        }
        data.lifestyle.lastProcessed = max(processed, now)
    }

    private static func randomAmount(_ lo: Decimal, _ hi: Decimal, using rng: inout SplitMix64) -> Decimal {
        guard lo != hi else { return lo }
        let a = min(lo.doubleValue, hi.doubleValue), b = max(lo.doubleValue, hi.doubleValue)
        let v = Double.random(in: a...b, using: &rng)
        // round to a plausible unit: 100 won for small amounts, 10,000 for large
        let unit: Double = abs(v) >= 1_000_000 ? 10_000 : (abs(v) >= 10_000 ? 1_000 : 100)
        return Decimal(double: (v / unit).rounded() * unit)
    }
}

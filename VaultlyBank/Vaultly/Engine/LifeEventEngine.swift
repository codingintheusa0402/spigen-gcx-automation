import Foundation

/// Things that just happen to you: 교통사고, 경찰 단속, 구속, 실직. Rolled once per simulated day from the lifestyle
/// seed (reproducible), at most one every `cooldownDays`. Illegal choices raise `heat`, which is what gets you arrested.
enum LifeEventKind: String, Codable, CaseIterable {
    case accident, police, jail, jobLoss
    var title: String {
        switch self {
        case .accident: return "교통사고"
        case .police: return "경찰 단속"
        case .jail: return "구속"
        case .jobLoss: return "실직"
        }
    }
    var emoji: String {
        switch self { case .accident: return "🚑"; case .police: return "🚔"; case .jail: return "⛓️"; case .jobLoss: return "📦" }
    }
}

struct LifeEventRecord: Codable, Equatable, Identifiable {
    var id = UUID()
    var date: Date
    var kind: LifeEventKind
    var title: String
    var detail: String
    var cash: Decimal
}

struct LifeEventState: Codable, Equatable {
    var records: [LifeEventRecord] = []
    var injuredUntil: Date? = nil        // character shows the hurt animation, 알바 closed
    var jailedUntil: Date? = nil         // salary off, 알바 closed, character in 구치소
    var lastEventDate: Date? = nil
    var heat: Double = 0                 // illegal history; decays daily
    var forced: LifeEventKind? = nil     // test hook: fire this on the next processed day

    func isInjured(at date: Date) -> Bool { injuredUntil.map { date < $0 } ?? false }
    func isJailed(at date: Date) -> Bool { jailedUntil.map { date < $0 } ?? false }
}

enum LifeEventEngine {
    static let cooldownDays = 30
    static let heatPerIllegalChoice = 1.0
    static let dailyHeatDecay = 0.997            // half-life ≈ 230 days

    /// Per-day probabilities (a 60-year life ≈ 22,000 days).
    static func chances(tier: WealthTier, game: GameState, life: LifeEventState) -> [LifeEventKind: Double] {
        var p: [LifeEventKind: Double] = [:]
        p[.accident] = 1.0 / 900                                                   // every ~2.5 years
        p[.police] = (1.0 / 450) * (1 + min(3, life.heat))                         // 범칙금 — more if you live on the edge
        p[.jail] = life.heat > 0 ? min(0.02, life.heat / 1500) : 0                 // only with an illegal history
        let hasJob = tier != .broke && tier != .tycoon && game.salaryPausedUntil == nil
        p[.jobLoss] = hasJob ? (1.0 / 1800) * (game.salaryMultiplier >= 1.4 ? 2.2 : 1) : 0   // rough high-salary places fire people
        return p
    }

    static func day(_ day: Date, now: Date, tier: WealthTier, accountID: UUID, rng: inout SplitMix64, data: inout BankData) -> [String] {
        guard !data.game.isEnded else { return [] }
        let cal = Calendar.current
        data.life.heat *= dailyHeatDecay
        if data.life.heat < 0.01 { data.life.heat = 0 }

        var kind: LifeEventKind? = nil
        var forced = false
        if let f = data.life.forced { kind = f; forced = true; data.life.forced = nil }
        else {
            if let last = data.life.lastEventDate, let gap = cal.dateComponents([.day], from: last, to: day).day, gap < cooldownDays { return [] }
            let roll = Double.random(in: 0..<1, using: &rng)
            var acc = 0.0
            for (k, p) in chances(tier: tier, game: data.game, life: data.life).sorted(by: { $0.key.rawValue < $1.key.rawValue }) {
                acc += p
                if roll < acc { kind = k; break }
            }
        }
        guard let k = kind else { return [] }
        guard var at = cal.date(bySettingHour: Int.random(in: 9...20, using: &rng), minute: Int.random(in: 0..<60, using: &rng), second: 7, of: day) else { return [] }
        if forced { at = min(now, max(at, data.lifestyle.lastProcessed.addingTimeInterval(1))) }   // test hook: always lands in the processed window
        guard at > data.lifestyle.lastProcessed, at <= now else { return [] }

        let scale = EventLibrary.tierScale[tier] ?? 1
        let brand = data.settings.brand
        func amount(_ lo: Double, _ hi: Double, factor: Double = 1) -> Decimal {
            let v = Double.random(in: lo...hi, using: &rng) * factor
            let unit: Double = v >= 1_000_000 ? 10_000 : 1_000
            return Decimal(double: (v / unit).rounded() * unit)
        }
        var record: LifeEventRecord
        switch k {
        case .accident:
            let bill = amount(150_000, 2_000_000, factor: pow(scale, 0.5) * data.game.medicalMultiplier)
            let hospital = ["서울대병원 응급실", "세브란스병원 응급의료센터", "삼성서울병원 응급실", "아산병원 응급실"][Int.random(in: 0..<4, using: &rng)]
            Ledger.apply(amount: -bill, to: accountID, date: at, counterparty: brand.resolve(hospital), memo: "교통사고 치료비", category: .transfer,
                         tag: SpendingCategory.medical.rawValue, data: &data)
            data.life.injuredUntil = cal.date(byAdding: .day, value: Int.random(in: 5...14, using: &rng), to: day)
            let insured = data.game.medicalMultiplier < 1
            record = LifeEventRecord(date: at, kind: k, title: "교통사고",
                                     detail: insured ? "출근길 접촉사고. 보험 덕에 치료비 일부만 냈습니다. 2주 정도 쉬어야 합니다."
                                                     : "출근길 접촉사고. 실손보험이 없어 치료비를 전부 냈습니다. 2주 정도 쉬어야 합니다.", cash: -bill)
        case .police:
            let fine = amount(40_000, 200_000, factor: 1 + min(2, data.life.heat))
            let what = ["속도위반 범칙금", "신호위반 범칙금", "불법주차 과태료", "음주 측정 거부 — 범칙금"][Int.random(in: 0..<(data.life.heat > 0 ? 4 : 3), using: &rng)]
            Ledger.apply(amount: -fine, to: accountID, date: at, counterparty: "경찰청 (교통)", memo: what, category: .transfer,
                         tag: SpendingCategory.tax.rawValue, data: &data)
            record = LifeEventRecord(date: at, kind: k, title: "경찰 단속", detail: "\(what) \(Money.format(fine, code: data.settings.currencyCode))을 냈습니다.", cash: -fine)
        case .jail:
            let months = Int.random(in: 3...12, using: &rng)
            let fine = amount(1_000_000, 20_000_000, factor: pow(scale, 0.5))
            Ledger.apply(amount: -fine, to: accountID, date: at, counterparty: "서울중앙지방법원", memo: "벌금 · 추징금", category: .transfer,
                         tag: SpendingCategory.tax.rawValue, data: &data)
            let until = cal.date(byAdding: .month, value: months, to: day)
            data.life.jailedUntil = until
            if let u = until { data.game.salaryPausedUntil = max(data.game.salaryPausedUntil ?? u, u) }
            data.life.heat = 0
            record = LifeEventRecord(date: at, kind: k, title: "구속 · 징역 \(months)개월",
                                     detail: "예전에 한 불법 일이 드러났습니다. \(months)개월 동안 수입이 끊기고 벌금 \(Money.format(fine, code: data.settings.currencyCode))을 냈습니다.", cash: -fine)
        case .jobLoss:
            let months = Int.random(in: 3...6, using: &rng)
            data.game.salaryPausedUntil = cal.date(byAdding: .month, value: months, to: day)
            let benefit: Decimal = min(1_980_000, max(1_200_000, amount(1_500_000, 1_980_000)))
            data.game.flows.append(ActiveFlow(questID: "life_jobloss", counterparty: "고용노동부", memo: "실업급여", amount: benefit, day: 10,
                                              category: .transfer, hour: 10, until: cal.date(byAdding: .month, value: 4, to: day)))
            Ledger.apply(amount: 0, to: accountID, date: at, counterparty: brand.resolve(data.game.employer ?? "회사"), memo: "권고사직 통보", category: .transfer,
                         tag: SpendingCategory.other.rawValue, data: &data)
            let why = data.game.salaryMultiplier >= 1.4 ? "성과 압박이 센 곳이었습니다. 하루아침에 권고사직." : "회사 사정으로 권고사직 통보를 받았습니다."
            record = LifeEventRecord(date: at, kind: k, title: "실직",
                                     detail: "\(why) \(months)개월 동안 급여가 없고, 4개월간 실업급여 \(Money.format(benefit, code: data.settings.currencyCode))이 나옵니다.", cash: 0)
        }
        data.life.records.append(record)
        if data.life.records.count > 200 { data.life.records.removeFirst(data.life.records.count - 200) }
        data.life.lastEventDate = day
        return ["⚠️ \(record.title) — \(record.detail)"]
    }
}

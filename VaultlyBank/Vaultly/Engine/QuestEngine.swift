import Foundation

/// Runs inside LifestyleEngine's day loop: offers quests, settles delayed outcomes, pays quest flows,
/// and ends the game at the end of the lifespan. Pure functions over BankData so it is testable.
enum QuestEngine {
    static func netWorth(_ d: BankData) -> Decimal {
        d.accounts.reduce(Decimal(0)) { $0 + ($1.kind == .loan ? -$1.balance : $1.balance) } + d.portfolio.investedKRW + d.portfolio.cashKRWTotal
    }

    static func eligible(_ q: Quest, tier: WealthTier, age: Int, game: GameState, netWorth: Decimal) -> Bool {
        guard q.tiers.contains(tier), age >= q.minAge, age <= q.maxAge else { return false }
        if q.once, game.answered(q.id) { return false }
        if let promo = q.promotion {
            guard GameState.nextTier(after: tier) == promo, netWorth >= (q.requiredNetWorth ?? 0) else { return false }
        }
        return true
    }

    static func pick(tier: WealthTier, age: Int, game: GameState, netWorth: Decimal, using rng: inout SplitMix64) -> Quest? {
        let pool = QuestLibrary.all.filter { eligible($0, tier: tier, age: age, game: game, netWorth: netWorth) }
        guard !pool.isEmpty else { return nil }
        let promos = pool.filter { $0.promotion != nil }
        if !promos.isEmpty, Double.random(in: 0..<1, using: &rng) < 0.6 {
            return promos[Int.random(in: 0..<promos.count, using: &rng)]
        }
        let normal = pool.filter { $0.promotion == nil }
        guard !normal.isEmpty else { return promos.first }
        return normal[Int.random(in: 0..<normal.count, using: &rng)]
    }

    /// One simulated day. Returns user-facing messages (outcomes landed, new quest, game over).
    static func day(_ day: Date, now: Date, tier: WealthTier, accountID: UUID, comps: DateComponents, lastDay: Int,
                    rng: inout SplitMix64, data: inout BankData) -> [String] {
        var msgs: [String] = []
        let cal = Calendar.current
        if data.game.birthDate == nil {
            data.game.birthDate = cal.date(byAdding: .year, value: -data.game.startAge, to: day)
        }
        guard !data.game.isEnded else { return msgs }
        let age = data.game.age(at: day)

        // 1. quest flows (보험료, 부업, 임대료, 상환 ...)
        data.game.flows.removeAll { f in f.until.map { $0 < day } ?? false }
        for (i, f) in data.game.flows.enumerated() where comps.day == min(f.day, lastDay) {
            let brand = data.settings.brand
            if let at = cal.date(bySettingHour: f.hour, minute: Int.random(in: 0..<60, using: &rng), second: 30 + i % 30, of: day),
               at > data.lifestyle.lastProcessed, at <= now {
                let cat = SpendingReport.category(of: Transaction(accountID: accountID, date: at, amount: f.amount, counterparty: f.counterparty, memo: f.memo, category: f.category, balanceAfter: 0))
                Ledger.apply(amount: f.amount, to: accountID, date: at, counterparty: brand.resolve(f.counterparty), memo: f.memo, category: f.category,
                             tag: cat.rawValue, data: &data)
            }
        }

        // 2. delayed outcomes
        let due = data.game.outcomes.filter { $0.fireDate <= day }
        for o in due {
            guard let at = cal.date(bySettingHour: 11, minute: Int.random(in: 0..<60, using: &rng), second: 0, of: day), at <= now else { continue }
            data.game.outcomes.removeAll { $0.id == o.id }
            let brand = data.settings.brand
            if o.cash != 0 {
                Ledger.apply(amount: o.cash, to: accountID, date: at, counterparty: brand.resolve(o.counterparty), memo: "[선택 결과] \(shortTitle(o.questID))",
                             category: .transfer, tag: (o.cash > 0 ? SpendingCategory.income : .other).rawValue, data: &data)
            }
            if o.fine > 0 {
                Ledger.apply(amount: -o.fine, to: accountID, date: at, counterparty: brand.resolve(o.counterparty), memo: "벌금·추징금", category: .transfer,
                             tag: SpendingCategory.tax.rawValue, data: &data)
            }
            if var f = o.flow {
                f.until = o.flowMonths.flatMap { cal.date(byAdding: .month, value: $0, to: at) }
                data.game.flows.append(f)
            }
            if o.salaryPause > 0 { data.game.salaryPausedUntil = cal.date(byAdding: .month, value: o.salaryPause, to: at) }
            if o.salaryMultiplier != 1 { data.game.salaryMultiplier *= o.salaryMultiplier }
            if let m = o.netWorthMultiplier, let idx = data.accounts.firstIndex(where: { $0.id == accountID }) {
                let bal = data.accounts[idx].balance
                if bal > 0 { Ledger.apply(amount: -(bal * Decimal(double: 1 - m)).rounded(), to: accountID, date: at, counterparty: brand.resolve(o.counterparty),
                                          memo: "[선택 결과] 재산 분할·압류", category: .transfer, tag: SpendingCategory.other.rawValue, data: &data) }
            }
            if let t = o.demoteTo { data.lifestyle.tier = t }
            if let ai = data.game.answers.firstIndex(where: { $0.id == o.answerID }) {
                data.game.answers[ai].outcome = o.text
                data.game.answers[ai].outcomeDate = at
            }
            msgs.append(o.text)
        }

        // 3. pending quest expiry
        if let exp = data.game.pendingExpires, data.game.pendingQuestID != nil, day > exp {
            data.game.pendingQuestID = nil; data.game.pendingSince = nil; data.game.pendingExpires = nil
            data.game.nextQuestDate = cal.date(byAdding: .day, value: 3, to: day)
        }

        // 4. offer a quest
        let nw = netWorth(data)
        if data.game.pendingQuestID == nil {
            if data.game.nextQuestDate == nil {
                scheduleNext(from: day, tier: tier, age: age, netWorth: nw, rng: &rng, data: &data)
            } else if let next = data.game.nextQuestDate, day >= next {
                var chosen: Quest? = data.game.nextQuestID.flatMap(QuestLibrary.quest).flatMap { eligible($0, tier: tier, age: age, game: data.game, netWorth: nw) ? $0 : nil }
                if chosen == nil { chosen = pick(tier: tier, age: age, game: data.game, netWorth: nw, using: &rng) }
                if let q = chosen {
                    data.game.pendingQuestID = q.id
                    data.game.pendingSince = day
                    data.game.pendingExpires = cal.date(byAdding: .day, value: GameState.questWindowDays, to: day)
                    data.game.nextQuestDate = nil; data.game.nextQuestID = nil
                    msgs.append("📩 \(q.title)")
                } else {
                    data.game.nextQuestDate = cal.date(byAdding: .day, value: 7, to: day)
                }
            }
        }

        // 5. end of life
        if age >= data.game.lifespanYears {
            data.game.endedAt = day
            data.game.pendingQuestID = nil
            msgs.append("🏁 \(data.game.lifespanYears)세. 인생이 끝났습니다. 결산을 확인하세요.")
        }
        return msgs
    }

    static func scheduleNext(from day: Date, tier: WealthTier, age: Int, netWorth: Decimal, rng: inout SplitMix64, data: inout BankData) {
        let gap = Int.random(in: GameState.questIntervalDays, using: &rng)
        data.game.nextQuestDate = Calendar.current.date(byAdding: .day, value: gap, to: day)
        data.game.nextQuestID = pick(tier: tier, age: age, game: data.game, netWorth: netWorth, using: &rng)?.id
    }

    static func shortTitle(_ questID: String) -> String {
        guard let q = QuestLibrary.quest(questID) else { return questID }
        return String(q.title.prefix(28))
    }

    /// Salary factor for this day: quest multiplier, 0 while paused.
    static func salaryFactor(_ game: GameState, day: Date) -> Double {
        if let until = game.salaryPausedUntil, day < until { return 0 }
        return game.salaryMultiplier
    }
}

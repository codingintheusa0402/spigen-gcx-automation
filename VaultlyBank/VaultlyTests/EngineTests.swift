import XCTest
@testable import Vaultly

final class EngineTests: XCTestCase {
    private let cal = Calendar.current
    private func date(_ y: Int, _ m: Int, _ d: Int, _ h: Int = 12) -> Date {
        cal.date(from: DateComponents(year: y, month: m, day: d, hour: h))!
    }

    // MARK: Money helpers

    func testDecimalRoundingAndFormatting() {
        XCTAssertEqual(Decimal(string: "1234.567")!.rounded(scale: 2), Decimal(string: "1234.57"))
        XCTAssertEqual(Decimal(double: 0.1 + 0.2), Decimal(string: "0.3"))
        let saved = Money.wonSuffix
        defer { Money.wonSuffix = saved }
        Money.wonSuffix = false
        XCTAssertEqual(Money.format(1_234_567, code: "KRW"), "₩1,234,567")
        XCTAssertTrue(Money.format(5000, code: "KRW", signed: true).hasPrefix("+"))
        Money.wonSuffix = true
        XCTAssertEqual(Money.format(1_234_567, code: "KRW"), "1,234,567원")
        XCTAssertEqual(Money.format(-7_890, code: "KRW"), "-7,890원")
        XCTAssertEqual(Money.format(27_456, code: "KRW", signed: true), "+27,456원")
        XCTAssertEqual(Money.format(Decimal(string: "2032.22")!, code: "USD"), "$2,032.22")
    }

    // MARK: Ledger / scheduled events

    func testScheduledEventAppliesAndRepeats() {
        var data = BankData.starter()
        let acc = data.accounts[0].id
        let start = data.accounts[0].balance
        let t0 = date(2026, 1, 1, 16)
        data.events = [ScheduledEvent(accountID: acc, counterparty: "엄마", amount: 5_000_000, memo: "용돈",
                                      nextFireDate: t0, repeatRule: .daily)]
        let fired = ScheduleEngine.processDue(now: date(2026, 1, 3, 17), data: &data)
        XCTAssertEqual(fired.count, 3)
        XCTAssertEqual(data.accounts[0].balance, start + 15_000_000)
        XCTAssertEqual(data.events[0].nextFireDate, date(2026, 1, 4, 16))
        XCTAssertTrue(data.events[0].isEnabled)
    }

    func testOneOffEventDisablesAfterFiring() {
        var data = BankData.starter()
        data.events = [ScheduledEvent(accountID: data.accounts[0].id, counterparty: "x", amount: 1, memo: "",
                                      nextFireDate: date(2026, 1, 1), repeatRule: .none)]
        _ = ScheduleEngine.processDue(now: date(2026, 2, 1), data: &data)
        XCTAssertFalse(data.events[0].isEnabled)
    }

    // MARK: Interest engine

    func testAnnuityPayment() {
        // 30,000,000 @ 4.2% / 36 months: 30M × 0.0035 / (1 − 1.0035⁻³⁶) ≈ 888,391
        let pay = InterestEngine.annuityPayment(principal: 30_000_000, annualRate: 0.042, months: 36)
        XCTAssertEqual(pay.doubleValue, 888_391, accuracy: 5)
        XCTAssertEqual(InterestEngine.annuityPayment(principal: 1200, annualRate: 0, months: 12), 100)
    }

    func testInstallmentSavingsProjection() {
        // 500,000 × 12 @ 5%: interest = 500k × 0.05 × (12+11+...+1)/12 = 162,500
        let terms = ProductTerms(annualRate: 0.05, termMonths: 12, startDate: Date(), principal: 0, monthlyAmount: 500_000,
                                 linkedAccountID: nil, lastAccrualDate: Date(), lastInstallmentDate: Date())
        let p = InterestEngine.projectedMaturity(terms: terms, kind: .installmentSavings)
        XCTAssertEqual(p.principal, 6_000_000)
        XCTAssertEqual(p.interest, 162_500)
    }

    func testLoanRepaymentsPulledMonthly() {
        var data = BankData.starter()
        let checking = data.accounts[0].id
        data.accounts[0].balance = 100_000_000
        let start = date(2026, 1, 15)
        var loan = Account(name: "loan", kind: .loan, number: "1", balance: 30_000_000)
        loan.terms = ProductTerms(annualRate: 0.042, termMonths: 36, startDate: start, principal: 30_000_000,
                                  monthlyAmount: InterestEngine.annuityPayment(principal: 30_000_000, annualRate: 0.042, months: 36),
                                  linkedAccountID: checking, lastAccrualDate: start, lastInstallmentDate: start)
        data.accounts.append(loan)
        InterestEngine.accrue(now: date(2026, 4, 1), data: &data, taxRate: 0.154)
        let debt = data.accounts[1].balance
        XCTAssertLessThan(debt, 30_000_000)
        XCTAssertGreaterThan(debt, 27_000_000)
        let repayments = data.transactions.filter { $0.accountID == checking && $0.category == .repayment }
        XCTAssertEqual(repayments.count, 2) // Feb 15, Mar 15
    }

    func testDepositMaturesAndPaysNetInterest() {
        var data = BankData.starter()
        let checking = data.accounts[0].id
        let start = date(2025, 1, 1)
        var dep = Account(name: "dep", kind: .deposit, number: "2", balance: 10_000_000)
        dep.terms = ProductTerms(annualRate: 0.04, termMonths: 12, startDate: start, principal: 10_000_000, monthlyAmount: 0,
                                 linkedAccountID: checking, lastAccrualDate: start, lastInstallmentDate: start)
        data.accounts.append(dep)
        let before = data.accounts[0].balance
        InterestEngine.accrue(now: date(2026, 1, 2), data: &data, taxRate: 0.154)
        XCTAssertTrue(data.accounts[1].terms!.isClosed)
        XCTAssertEqual(data.accounts[1].balance, 0)
        let gained = data.accounts[0].balance - before
        // ≈ 10,000,000 + 400,000 × (1 − 0.154) = 10,338,400 (daily accrual, ±1 day)
        XCTAssertEqual(gained.doubleValue, 10_338_400, accuracy: 3_000)
    }

    // MARK: Clock

    func testClockSpeedAndMonotonicity() {
        var clock = SimClock()
        let real0 = Date()
        clock.anchorReal = real0; clock.anchorVirtual = real0
        clock.setSpeed(1200, real: real0)
        let later = real0.addingTimeInterval(3600)
        XCTAssertEqual(clock.now(real: later).timeIntervalSince(real0), 3600 * 1200, accuracy: 1)
        clock.setSpeed(1, real: later)
        XCTAssertEqual(clock.now(real: later).timeIntervalSince(real0), 3600 * 1200, accuracy: 1)
        XCTAssertFalse(clock.isRealTime)
        // never goes backwards
        XCTAssertEqual(clock.now(real: real0), clock.now(real: later))
    }

    // MARK: Lifestyle

    func testEventLibraryHasOverThousandVariationsPerTier() {
        for tier in WealthTier.allCases {
            XCTAssertGreaterThanOrEqual(EventLibrary.variationCount(for: tier), 1000, "\(tier)")
        }
    }

    func testLifestyleIsDeterministicForSameSeed() {
        func run() -> [Transaction] {
            var data = BankData.starter()
            data.lifestyle = LifestyleState(tier: .middle, accountID: data.accounts[0].id, debtAccountID: nil,
                                            tierStart: date(2026, 1, 1), lastProcessed: date(2026, 1, 1), seed: 42)
            LifestyleEngine.process(now: date(2026, 3, 1), data: &data)
            return data.transactions
        }
        let a = run(), b = run()
        XCTAssertEqual(a.map(\.amount), b.map(\.amount))
        XCTAssertEqual(a.map(\.memo), b.map(\.memo))
        XCTAssertGreaterThan(a.count, 60)
    }

    func testEverydayMerchantsStayRealisticForTycoon() {
        // 재벌이어도 GS25·맥도날드 결제가 수백만 원이 될 수는 없다 — scaling은 .fixed / .capped 로 묶여 있어야 한다.
        XCTAssertEqual(LifeEventSpec.Scaling.fixed.factor(90), 1)
        XCTAssertEqual(LifeEventSpec.Scaling.capped(1.5).factor(90), 1.5)
        XCTAssertEqual(LifeEventSpec.Scaling.capped(1.5).factor(0.18), 0.6, accuracy: 0.0001)   // 거지도 밥값은 반값 이하로 안 내려감
        XCTAssertEqual(LifeEventSpec.Scaling.full.factor(90), 90)

        var data = BankData.starter()
        data.lifestyle = LifestyleState(tier: .tycoon, accountID: data.accounts[0].id, debtAccountID: nil,
                                        tierStart: date(2026, 1, 1), lastProcessed: date(2026, 1, 1), seed: 7)
        LifestyleEngine.process(now: date(2026, 4, 1), data: &data)
        let everyday = ["GS25", "CU ", "세븐일레븐", "이마트24", "맥도날드", "스타벅스", "김밥천국", "맘스터치", "카카오T", "우버", "타다"]
        let hits = data.transactions.filter { tx in everyday.contains { tx.counterparty.contains($0) } }
        XCTAssertGreaterThan(hits.count, 20, "tycoon still buys coffee and takes taxis")
        for tx in hits {
            XCTAssertLessThanOrEqual(abs(tx.amount), 150_000, "\(tx.counterparty) \(tx.amount)")
        }
        XCTAssertFalse(data.transactions.contains { $0.counterparty == "모범택시" || $0.counterparty == "편의점" || $0.counterparty == "맥도날드" })
    }

    // MARK: 챌린지 · 목표 · 리포트 · 호환성

    func testOldDocumentWithoutNewKeysStillDecodes() throws {
        // A bank.json written before challenge/goals/achievements existed must load, not wipe the user's data.
        var old = BankData.starter()
        old.accounts[0].balance = 777
        let enc = JSONEncoder(); enc.dateEncodingStrategy = .iso8601
        var json = try JSONSerialization.jsonObject(with: enc.encode(old)) as! [String: Any]
        json.removeValue(forKey: "challenge"); json.removeValue(forKey: "goals"); json.removeValue(forKey: "achievements")
        let raw = try JSONSerialization.data(withJSONObject: json)
        let dec = JSONDecoder(); dec.dateDecodingStrategy = .iso8601
        let back = try dec.decode(BankData.self, from: raw)
        XCTAssertEqual(back.accounts[0].balance, 777)
        XCTAssertTrue(back.goals.isEmpty)
        XCTAssertNil(back.challenge.accountID)
    }

    func testChallengeStreakCountsConsecutiveDays() {
        var c = ChallengeState()
        let today = date(2026, 5, 10)
        for d in [8, 9, 10] { c.items.append(ResistedPurchase(date: date(2026, 5, d), title: "x", amount: 5_000, kind: .coffee)) }
        c.items.append(ResistedPurchase(date: date(2026, 5, 3), title: "y", amount: 20_000, kind: .delivery))
        XCTAssertEqual(c.streak(now: today), 3)
        XCTAssertEqual(c.streak(now: date(2026, 5, 11)), 3, "yesterday still counts")
        XCTAssertEqual(c.streak(now: date(2026, 5, 13)), 0)
        XCTAssertEqual(c.bestStreak(), 3)
        XCTAssertEqual(c.total(inMonthOf: today), 35_000)
    }

    @MainActor func testLogResistDepositsIntoChallengeAccount() {
        let store = BankStore()
        store.resetAll()
        let before = store.data.accounts.count
        let item = store.logResist(kind: .delivery, title: "", amount: 24_000)
        XCTAssertEqual(item?.title, ResistKind.delivery.placeholder)
        XCTAssertEqual(store.data.accounts.count, before + 1)
        XCTAssertEqual(store.challengeAccount?.balance, 24_000)
        XCTAssertNotNil(store.data.achievements[AchievementID.firstResist.rawValue])
        store.deleteResist(item!.id)
        XCTAssertEqual(store.challengeAccount?.balance, 0)
        store.resetAll()
    }

    func testGoalAutoTransferMovesMoneyBetweenOwnAccounts() {
        var data = BankData.starter()
        let from = data.accounts[0].id
        let goalAcc = Account(name: "✈️ 제주", kind: .deposit, number: "1", balance: 0)
        data.accounts.append(goalAcc)
        data.events = [ScheduledEvent(accountID: from, counterparty: goalAcc.name, amount: -300_000, memo: "목표 자동이체",
                                      nextFireDate: date(2026, 1, 25), repeatRule: .monthly, toAccountID: goalAcc.id)]
        ScheduleEngine.processDue(now: date(2026, 3, 1), data: &data)
        XCTAssertEqual(data.accounts[0].balance, 1_250_000 - 600_000)
        XCTAssertEqual(data.accounts[1].balance, 600_000)
        XCTAssertTrue(data.transactions.filter { $0.tag == SpendingCategory.transfer.rawValue }.count == 4)
        // the report must not count internal transfers as spending
        let r = SpendingReport.build(transactions: data.transactions, month: date(2026, 2, 1))
        XCTAssertEqual(r.spent, 0)
    }

    func testSpendingReportGroupsByCategoryAndComparesMonths() {
        var data = BankData.starter()
        let acc = data.accounts[0].id
        Ledger.apply(amount: -12_000, to: acc, date: date(2026, 4, 3), counterparty: "(주)맥도날드 삼성중앙역점", memo: "점심", category: .transfer, tag: "food", data: &data)
        Ledger.apply(amount: -5_500, to: acc, date: date(2026, 4, 4), counterparty: "스타벅스 강남R점", memo: "커피", category: .transfer, data: &data)   // keyword fallback
        Ledger.apply(amount: -36_000, to: acc, date: date(2026, 4, 5), counterparty: "KT", memo: "인터넷 요금", category: .autoTransfer, data: &data)
        Ledger.apply(amount: 2_800_000, to: acc, date: date(2026, 4, 25), counterparty: "(주)대성물류", memo: "급여", category: .transfer, data: &data)
        Ledger.apply(amount: -100_000, to: acc, date: date(2026, 3, 9), counterparty: "GS25 서초점", memo: "편의점", category: .transfer, tag: "mart", data: &data)
        let r = SpendingReport.build(transactions: data.transactions, month: date(2026, 4, 1))
        XCTAssertEqual(r.spent, 53_500)
        XCTAssertEqual(r.income, 2_800_000)
        XCTAssertEqual(r.slices.map(\.category), [.housing, .food, .cafe])
        XCTAssertEqual(r.merchants.first?.name, "KT")
        XCTAssertEqual(r.previousSpent, 100_000)
        XCTAssertEqual(r.changeRatio!, -0.465, accuracy: 0.001)
        XCTAssertEqual(SpendingCategory.forSpec("taxi_premium"), .transport)
        XCTAssertEqual(SpendingCategory.forSpec("ins_car"), .insurance)
    }

    @MainActor func testScenarioAppliesTierAndSchedulesEvents() {
        let store = BankStore()
        store.resetAll()
        store.applyScenario(.payday)
        XCTAssertEqual(store.lifestyle.tier, .working)
        XCTAssertEqual(store.data.events.count, 3)
        XCTAssertEqual(store.data.events.first?.counterparty, "(주)대성물류")
        XCTAssertFalse(store.data.events.contains { $0.counterparty.contains("{card}") }, "placeholders resolved")
        store.resetAll()
    }

    func testBankLineMatchesRealPushFormat() {
        let saved = Money.wonSuffix; Money.wonSuffix = true; defer { Money.wonSuffix = saved }
        let at = Calendar.current.date(bySettingHour: 15, minute: 40, second: 0, of: date(2028, 3, 9))!
        let line = NotificationService.bankLine(bank: "우리은행", date: at, incoming: true, amount: 31_440_000,
                                                counterparty: "토스증권", balance: 123_760_684_504, currency: "KRW")
        XCTAssertEqual(line, "[우리은행] 03/09 15:40 입금 31,440,000원 토스증권 잔액 123,760,684,504원")
    }

    // MARK: 인생 게임 (quests · tiers)

    func testQuestLibraryIsLargeAndWellFormed() {
        let all = QuestLibrary.all
        XCTAssertGreaterThanOrEqual(all.count, 300)
        XCTAssertEqual(Set(all.map(\.id)).count, all.count, "ids unique")
        for q in all { XCTAssertGreaterThanOrEqual(q.options.count, 2, q.id); XCTAssertFalse(q.title.isEmpty, q.id) }
        for t in WealthTier.allCases { XCTAssertGreaterThanOrEqual(all.filter { $0.tiers.contains(t) }.count, 50, "\(t)") }
        for t in [WealthTier.working, .middle, .rich, .tycoon] { XCTAssertTrue(all.contains { $0.promotion == t }, "promotion quest to \(t)") }
        XCTAssertTrue(all.contains { $0.title.contains("삼성전자") && $0.title.contains("이직") })
        XCTAssertTrue(all.contains { $0.title.contains("종신보험") })
        XCTAssertTrue(all.contains { $0.id.hasPrefix("illegal_") && $0.options[0].effect.gamble?.loseSalaryPause ?? 0 > 0 }, "illegal job can end in prison")
    }

    func testQuestIsOfferedWithinTwoWeeksAndExpires() {
        var data = BankData.starter()
        data.lifestyle = LifestyleState(tier: .working, accountID: data.accounts[0].id, debtAccountID: nil, tierStart: date(2026, 1, 1), lastProcessed: date(2026, 1, 1), seed: 9)
        LifestyleEngine.process(now: date(2026, 1, 20), data: &data)
        XCTAssertNotNil(data.game.pendingQuestID, "a quest should be offered within 14 days")
        XCTAssertNotNil(data.game.birthDate)
        XCTAssertEqual(data.game.age(at: date(2026, 1, 20)), 22)
        let q = QuestLibrary.quest(data.game.pendingQuestID!)!
        XCTAssertTrue(q.tiers.contains(.working))
        XCTAssertNil(q.promotion, "promotion quests need net worth first")
        LifestyleEngine.process(now: date(2026, 3, 1), data: &data)   // 14-day window passed → expires, new one arrives
        XCTAssertNotNil(data.game.pendingQuestID)
    }

    @MainActor func testAnsweringInsuranceQuestAddsPremiumFlowAndLowersMedicalBills() {
        let store = BankStore(); store.resetAll()
        store.forcePendingQuest("ins_01")                                  // 삼성화재 실손
        let before = store.data.accounts[0].balance
        let r = store.answerQuest(0)
        XCTAssertNotNil(r)
        XCTAssertEqual(store.game.flows.count, 1)
        XCTAssertEqual(store.game.flows[0].amount, -95_000)
        XCTAssertEqual(store.game.medicalMultiplier, 0.35, accuracy: 0.001)
        XCTAssertEqual(store.data.accounts[0].balance, before)
        XCTAssertNil(store.pendingQuest)
        XCTAssertNotNil(store.game.nextQuestDate)
        // declining raises hospital bills instead
        store.resetAll(); store.forcePendingQuest("ins_01"); store.answerQuest(1)
        XCTAssertEqual(store.game.medicalMultiplier, 1.8, accuracy: 0.001)
        store.resetAll()
    }

    @MainActor func testJobChangeRaisesSalaryAndCashQuestMovesMoney() {
        let store = BankStore(); store.resetAll()
        store.forcePendingQuest("career_001")
        let q = store.pendingQuest!
        let mult = q.options[0].effect.salaryMultiplier
        store.answerQuest(0)
        XCTAssertEqual(store.game.salaryMultiplier, mult, accuracy: 0.001)
        XCTAssertNotNil(store.game.employer)
        store.forcePendingQuest("lux_bag")
        let before = store.data.accounts[0].balance
        store.answerQuest(0)
        XCTAssertEqual(store.data.accounts[0].balance, before - 14_000_000)
        XCTAssertEqual(store.game.answers.count, 2)
        store.resetAll()
    }

    @MainActor func testGambleResolvesMonthsLaterViaEngine() {
        let store = BankStore(); store.resetAll()
        let id = QuestLibrary.all.first { $0.id.hasPrefix("inv_bitcoin") }!.id
        store.forcePendingQuest(id)
        store.answerQuest(0)
        XCTAssertEqual(store.game.outcomes.count, 1)
        XCTAssertNil(store.game.answers[0].outcome)
        var d = store.data
        let later = Calendar.current.date(byAdding: .month, value: 14, to: d.clock.now())!
        LifestyleEngine.process(now: later, data: &d)
        XCTAssertTrue(d.game.outcomes.isEmpty, "outcome settled")
        XCTAssertNotNil(d.game.answers[0].outcome)
        XCTAssertTrue(d.transactions.contains { $0.memo.hasPrefix("[선택 결과]") })
        store.resetAll()
    }

    @MainActor func testTiersAreLockedUntilPromotionQuest() {
        let store = BankStore(); store.resetAll()
        let acc = store.data.accounts[0].id
        store.applyTier(.middle, accountID: acc, resetBalance: false, backfillDays: 0)
        XCTAssertNil(store.lifestyle.tier, "중산층 is locked at start")
        store.applyTier(.working, accountID: acc, resetBalance: false, backfillDays: 0)
        XCTAssertEqual(store.lifestyle.tier, .working)
        // promotion quest cannot be picked below the threshold
        var rng = SplitMix64(seed: 1)
        let pick = QuestEngine.pick(tier: .working, age: 30, game: store.game, netWorth: 1_000, using: &rng)
        XCTAssertNil(pick?.promotion)
        XCTAssertTrue(QuestEngine.eligible(QuestLibrary.quest("up_middle_1")!, tier: .working, age: 30, game: store.game, netWorth: 60_000_000))
        store.forcePendingQuest("up_middle_1")
        store.answerQuest(0)
        XCTAssertEqual(store.lifestyle.tier, .middle)
        XCTAssertTrue(store.game.isUnlocked(.middle))
        XCTAssertFalse(store.game.isUnlocked(.rich))
        store.resetAll()
    }

    func testLifeEndsAtEightyAndMaxSpeedFitsTwelveHours() {
        XCTAssertEqual(TimeSpeed.lifeIn12h.rawValue * 12 * 3600 / (365.25 * 86400), 60, accuracy: 0.1, "12 real hours = 60 simulated years")
        var data = BankData.starter()
        data.lifestyle = LifestyleState(tier: .working, accountID: data.accounts[0].id, debtAccountID: nil, tierStart: date(2026, 1, 1), lastProcessed: date(2026, 1, 1), seed: 3)
        data.game.birthDate = Calendar.current.date(byAdding: .year, value: -80, to: date(2026, 1, 10))
        LifestyleEngine.process(now: date(2026, 2, 1), data: &data)
        XCTAssertTrue(data.game.isEnded)
        let count = data.transactions.count
        LifestyleEngine.process(now: date(2026, 6, 1), data: &data)
        XCTAssertEqual(data.transactions.count, count, "nothing happens after the ending")
    }

    func testSalaryRisesWithYears() {
        var data = BankData.starter()
        data.accounts[0].balance = 0
        data.lifestyle = LifestyleState(tier: .middle, accountID: data.accounts[0].id, debtAccountID: nil,
                                        tierStart: date(2026, 1, 1), lastProcessed: date(2026, 1, 1), seed: 7)
        // process in 120-day chunks (engine budget) over ~3 years
        var now = date(2026, 1, 1)
        let end = date(2029, 2, 1)
        var guardCount = 0
        while now < end, guardCount < 60 {
            guardCount += 1
            now = cal.date(byAdding: .day, value: 100, to: now)!
            LifestyleEngine.process(now: min(now, end), data: &data)
            // drain any leftover budget for this window
            LifestyleEngine.process(now: min(now, end), data: &data)
        }
        let salaries = data.transactions.filter { $0.memo.hasPrefix("급여") }.sorted { $0.date < $1.date }
        XCTAssertGreaterThan(salaries.count, 30)
        XCTAssertGreaterThan(salaries.last!.amount, salaries.first!.amount)
        XCTAssertTrue(salaries.last!.memo.contains("년차"))
    }

    func testBrokeTierDebtOnlyGrows() {
        var data = BankData.starter()
        let checking = data.accounts[0].id
        data.accounts[0].balance = 13_240
        let debt = Account(name: "OK캐피탈", kind: .loan, number: "9", balance: 4_800_000)
        data.accounts.append(debt)
        data.lifestyle = LifestyleState(tier: .broke, accountID: checking, debtAccountID: debt.id,
                                        tierStart: date(2026, 1, 1), lastProcessed: date(2026, 1, 1), seed: 3)
        var now = date(2026, 1, 1)
        for _ in 0..<8 {
            now = cal.date(byAdding: .day, value: 90, to: now)!
            LifestyleEngine.process(now: now, data: &data)
            LifestyleEngine.process(now: now, data: &data)
        }
        XCTAssertGreaterThan(data.accounts[1].balance, Decimal(7_200_000))   // ~2 years → much more debt
        XCTAssertGreaterThanOrEqual(data.accounts[0].balance, 0)           // auto top-up keeps checking ≥ 0
    }

    func testLifestyleBudgetDoesNotSkipDays() {
        var data = BankData.starter()
        data.lifestyle = LifestyleState(tier: .working, accountID: data.accounts[0].id, debtAccountID: nil,
                                        tierStart: date(2026, 1, 1), lastProcessed: date(2026, 1, 1), seed: 11)
        // 400 days in one call exceeds the per-call budget; repeated calls must cover all months
        for _ in 0..<5 { LifestyleEngine.process(now: date(2027, 2, 5), data: &data) }
        let salaries = data.transactions.filter { $0.memo.hasPrefix("급여") }
        XCTAssertEqual(salaries.count, 13, "one salary per month from Jan 2026 to Jan 2027")
    }

    // MARK: Trading

    func testTradeUpdatesCashAndAvgCost() {
        var p = Portfolio()
        p.cashKRW = 1_000_000
        p.holdings = [Holding(symbol: "A", name: "A", market: .domestic, shares: 10, avgCost: 100, price: 100)]
        // emulate BankStore.trade logic through a store on the main actor
        let exp = expectation(description: "trade")
        Task { @MainActor in
            let store = BankStore()
            store.updatePortfolio { $0 = p }
            let id = store.portfolio.holdings[0].id
            XCTAssertTrue(store.trade(holdingID: id, buy: true, shares: 10, price: 200))
            XCTAssertEqual(store.portfolio.cashKRW, 1_000_000 - 2_000)
            XCTAssertEqual(store.portfolio.holdings[0].avgCost, 150)
            XCTAssertFalse(store.trade(holdingID: id, buy: true, shares: 100_000, price: 200)) // insufficient cash
            XCTAssertTrue(store.trade(holdingID: id, buy: false, shares: 5, price: 300))
            XCTAssertEqual(store.portfolio.holdings[0].shares, 15)
            exp.fulfill()
        }
        wait(for: [exp], timeout: 5)
    }
}

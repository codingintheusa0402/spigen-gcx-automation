import XCTest
@testable import Vaultly

/// 공휴일 · 인생 이벤트 · 스쿠터 레이스 · 블랙잭 — pure-logic tests (no UI).
final class LifeAndGamesTests: XCTestCase {
    private var cal: Calendar { var c = Calendar(identifier: .gregorian); c.timeZone = TimeZone(identifier: "Asia/Seoul")!; return c }
    private func d(_ y: Int, _ m: Int, _ day: Int, _ h: Int = 12) -> Date { cal.date(from: DateComponents(year: y, month: m, day: day, hour: h))! }
    private func sod(_ y: Int, _ m: Int, _ day: Int) -> Date { cal.startOfDay(for: d(y, m, day)) }

    // MARK: 공휴일

    func testFixedAndLunarHolidays2026() {
        let names = Dictionary(KoreanHolidays.holidays(in: 2026, calendar: cal).map { ($0.date, $0.name) }, uniquingKeysWith: { a, _ in a })
        XCTAssertEqual(names[sod(2026, 1, 1)], "신정")
        XCTAssertEqual(names[sod(2026, 2, 17)], "설날")             // 음력 1월 1일
        XCTAssertEqual(names[sod(2026, 2, 16)], "설날 연휴")
        XCTAssertEqual(names[sod(2026, 2, 18)], "설날 연휴")
        XCTAssertEqual(names[sod(2026, 5, 24)], "부처님오신날")      // 음력 4월 8일
        XCTAssertEqual(names[sod(2026, 9, 25)], "추석")             // 음력 8월 15일
        XCTAssertEqual(names[sod(2026, 10, 3)], "개천절")
        XCTAssertEqual(names[sod(2026, 12, 25)], "성탄절")
        XCTAssertNil(names[sod(2026, 7, 17)], "제헌절 is not a public holiday")
    }

    func testSubstituteHolidays2026() {
        let subs = KoreanHolidays.holidays(in: 2026, calendar: cal).filter { $0.substitute }
        let dates = Set(subs.map { $0.date })
        XCTAssertTrue(dates.contains(sod(2026, 3, 2)), "삼일절 on Sunday → Monday")
        XCTAssertTrue(dates.contains(sod(2026, 5, 25)), "부처님오신날 on Sunday → Monday")
        XCTAssertTrue(dates.contains(sod(2026, 8, 17)), "광복절 on Saturday → Monday")
        XCTAssertTrue(dates.contains(sod(2026, 10, 5)), "개천절 on Saturday → Monday")
        XCTAssertFalse(dates.contains(sod(2026, 6, 8)), "현충일 never gets a substitute")
        XCTAssertFalse(dates.contains(sod(2026, 9, 28)), "추석 연휴 on Saturday only → no substitute (Sunday rule)")
    }

    func testWorkdayAndNextHoliday() {
        XCTAssertFalse(KoreanHolidays.isWorkday(d(2026, 10, 3), calendar: cal))   // Saturday + 개천절
        XCTAssertFalse(KoreanHolidays.isWorkday(d(2026, 10, 9), calendar: cal))   // 한글날 (Friday)
        XCTAssertTrue(KoreanHolidays.isWorkday(d(2026, 10, 6), calendar: cal))
        XCTAssertEqual(KoreanHolidays.next(after: d(2026, 10, 6), calendar: cal)?.name, "한글날")
        XCTAssertEqual(KoreanHolidays.next(after: d(2026, 12, 26), calendar: cal)?.name, "신정")   // rolls into 2027
    }

    func testRoutineTreatsHolidayAsDayOff() {
        let holiday = DailyRoutine.slot(tier: .working, hour: 10, workday: KoreanHolidays.isWorkday(d(2026, 10, 9), calendar: cal), variant: 1)
        let weekday = DailyRoutine.slot(tier: .working, hour: 10, workday: KoreanHolidays.isWorkday(d(2026, 10, 6), calendar: cal), variant: 1)   // variant 0 works in a 매장
        XCTAssertNotEqual(holiday.activity, "office")
        XCTAssertEqual(weekday.activity, "office")
    }

    // MARK: 인생 이벤트

    private func workingLife(at now: Date) -> BankData {
        var data = BankData.starter()
        let acc = data.accounts[0].id
        Ledger.setBalance(5_000_000, of: acc, date: now.addingTimeInterval(-86_400 * 2), data: &data)
        data.lifestyle.tier = .working; data.lifestyle.accountID = acc
        data.lifestyle.tierStart = now.addingTimeInterval(-86_400 * 30)
        data.lifestyle.lastProcessed = now.addingTimeInterval(-3_600 * 20)
        data.game.birthDate = cal.date(byAdding: .year, value: -30, to: now)
        data.game.nextQuestDate = cal.date(byAdding: .day, value: 10, to: now)
        return data
    }

    func testForcedAccidentChargesHospitalAndInjures() {
        let now = d(2026, 10, 6, 21)
        var data = workingLife(at: now)
        let before = data.accounts[0].balance
        data.life.forced = .accident
        let msgs = LifestyleEngine.process(now: now, data: &data)
        XCTAssertEqual(data.life.records.count, 1)
        XCTAssertEqual(data.life.records.first?.kind, .accident)
        XCTAssertLessThan(data.accounts[0].balance, before, "hospital bill was charged")
        XCTAssertTrue(data.life.isInjured(at: now))
        XCTAssertTrue(msgs.contains { $0.contains("교통사고") })
        XCTAssertNil(data.life.forced)
    }

    func testNoInsuranceMakesAccidentsCostMore() {
        let now = d(2026, 10, 6, 21)
        var insured = workingLife(at: now); insured.game.medicalMultiplier = 0.4; insured.life.forced = .accident
        var bare = workingLife(at: now); bare.game.medicalMultiplier = 1.8; bare.life.forced = .accident
        LifestyleEngine.process(now: now, data: &insured)
        LifestyleEngine.process(now: now, data: &bare)
        // same seed → same base roll; the multiplier is the only difference
        XCTAssertLessThan(-(insured.life.records[0].cash), -(bare.life.records[0].cash))
    }

    func testForcedJailStopsSalaryAndFines() {
        let now = d(2026, 10, 6, 21)
        var data = workingLife(at: now)
        data.life.heat = 3
        data.life.forced = .jail
        LifestyleEngine.process(now: now, data: &data)
        XCTAssertEqual(data.life.records.first?.kind, .jail)
        XCTAssertTrue(data.life.isJailed(at: now))
        XCTAssertNotNil(data.game.salaryPausedUntil)
        XCTAssertEqual(QuestEngine.salaryFactor(data.game, day: now.addingTimeInterval(86_400 * 30)), 0, "no salary while in jail")
        XCTAssertEqual(data.life.heat, 0, "served the sentence")
        XCTAssertLessThan(data.accounts[0].balance, 5_000_000, "fine charged")
    }

    func testForcedJobLossAddsUnemploymentBenefit() {
        let now = d(2026, 10, 6, 21)
        var data = workingLife(at: now)
        data.life.forced = .jobLoss
        LifestyleEngine.process(now: now, data: &data)
        XCTAssertEqual(data.life.records.first?.kind, .jobLoss)
        XCTAssertTrue(data.game.flows.contains { $0.memo == "실업급여" && $0.amount > 0 })
        XCTAssertNotNil(data.game.salaryPausedUntil)
    }

    func testEventsAreRareAndRespectCooldown() {
        let now = d(2026, 10, 6, 21)
        var data = workingLife(at: now)
        let p = LifeEventEngine.chances(tier: .working, game: data.game, life: data.life)
        XCTAssertEqual(p[.jail], 0, "no illegal history → no arrest")
        XCTAssertLessThan(p.values.reduce(0, +), 0.01)
        data.life.lastEventDate = cal.date(byAdding: .day, value: -3, to: now)
        var rng = SplitMix64(seed: 1)
        XCTAssertTrue(LifeEventEngine.day(cal.startOfDay(for: now), now: now, tier: .working, accountID: data.accounts[0].id, rng: &rng, data: &data).isEmpty)
        // heat from illegal choices makes jail possible and decays
        data.life.heat = 2
        XCTAssertGreaterThan(LifeEventEngine.chances(tier: .working, game: data.game, life: data.life)[.jail]!, 0)
    }

    // MARK: 스쿠터 레이스

    func testScooterShiftEndsAndPaysPerDelivery() {
        var g = ScooterGame(seed: 7)
        XCTAssertEqual(g.kind, .scooter)
        var t = 0.0
        while !g.isOver && t < 70 { g.steer = t.truncatingRemainder(dividingBy: 2) < 1 ? -1 : 1; g.tick(1.0 / 30); t += 1.0 / 30 }
        XCTAssertTrue(g.isOver)
        let r = g.result()
        XCTAssertEqual(r.kind, .scooter)
        let expected = ScooterGame.feePerDelivery * Decimal(g.delivered)
        if g.accident { XCTAssertLessThanOrEqual(r.payout, expected - ScooterGame.erBill + expected) }
        else { XCTAssertGreaterThanOrEqual(r.payout, expected) }
        XCTAssertFalse(r.lines.isEmpty)
    }

    func testScooterIsDeterministicPerSeed() {
        var a = ScooterGame(seed: 99), b = ScooterGame(seed: 99)
        for _ in 0..<300 { a.tick(0.05); b.tick(0.05) }
        XCTAssertEqual(a.markers, b.markers)
        XCTAssertEqual(a.crashes, b.crashes)
    }

    // MARK: 블랙잭

    func testBlackjackTotalsCountAcesSoftly() {
        func c(_ r: Int) -> CardGame.Card { CardGame.Card(rank: r, suit: 0, serial: r) }
        XCTAssertEqual(CardGame.total([c(1), c(13)]), 21)
        XCTAssertEqual(CardGame.total([c(1), c(1), c(9)]), 21)
        XCTAssertEqual(CardGame.total([c(1), c(9), c(9)]), 19)
        XCTAssertEqual(CardGame.total([c(10), c(5), c(9)]), 24)
    }

    func testBlackjackChipsAccounting() {
        var g = CardGame(seed: 3, stake: 100_000)
        XCTAssertEqual(g.bet, 10_000)
        XCTAssertEqual(g.chips, 100_000)
        g.deal()
        XCTAssertEqual(g.hands, 1)
        if g.phase == .player { g.stand() }
        let settled: Bool = { if case .settled = g.phase { return true } else { return false } }()
        XCTAssertTrue(settled)
        let allowed: Set<Decimal> = [90_000, 100_000, 110_000, 115_000]
        XCTAssertTrue(allowed.contains(g.chips), "chips after one hand must be lose/push/win/blackjack: \(g.chips)")
        XCTAssertEqual(g.result().payout, g.chips - 100_000)
        XCTAssertEqual(g.result().kind, .cards)
    }

    func testBlackjackSessionEndsWhenBroke() {
        var g = CardGame(seed: 5, stake: 20_000)   // bet 2,000; keep playing until the clock or the chips run out
        var rounds = 0
        while !g.isOver && rounds < 500 {
            g.deal()
            if g.phase == .player { g.hit(); if g.phase == .player { g.stand() } }
            g.tick(2.0)      // let the settle timer elapse
            rounds += 1
        }
        XCTAssertTrue(g.isOver)
        XCTAssertGreaterThan(g.hands, 1)
    }

    func testCardStakesByTier() {
        XCTAssertEqual(CardGame.stakes(for: 0), [10_000, 50_000], "every tier can buy in")
        XCTAssertEqual(CardGame.stakes(for: 1), [50_000, 100_000])
        for level in 0...4 { XCTAssertNil(JobKind.cards.lockedReason(for: level)); XCTAssertNil(JobKind.dayTrading.lockedReason(for: level)) }
        XCTAssertFalse(DayTradingGame.stakes(for: 0).isEmpty, "단타 open to 거지 too")
        XCTAssertTrue(JobKind.cards.isEntertainment)
        XCTAssertFalse(JobKind.scooter.isEntertainment)
    }
}

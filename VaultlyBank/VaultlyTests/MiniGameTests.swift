import XCTest
#if canImport(MiniGames)
@testable import MiniGames      // Linux SwiftPM harness for the pure game core
#else
@testable import Vaultly
#endif

/// Each part-time job is played start-to-finish by a skilled bot (should earn a good grade) and a careless or idle
/// bot (should not), over many seeds. Money rules are checked separately: wages, damage, and trading losses.
final class MiniGameTests: XCTestCase {
    let dt = 0.1
    let seeds: [UInt64] = Array(1...25)

    // MARK: 편의점
    func testConvenienceSkilledAndIdle() {
        var grades: [Grade] = []
        for s in seeds {
            var g = ConvenienceGame(seed: s)
            var steps = 0
            while !g.isOver && steps < 10_000 {
                steps += 1
                if steps % 3 == 0, let c = g.current {           // a human-ish pace: one action per 0.3 s
                    if g.phase == .scanning, let i = c.scanned.firstIndex(of: false) { g.scan(i) }
                    else if g.phase == .change {
                        let need = c.change - g.changeDrawer
                        if need == 0 { g.giveChange() }
                        else if let d = ConvenienceGame.denominations.first(where: { $0 <= need }) { g.addChange(d) }
                        else { g.giveChange() }   // change below 100원 is impossible with these prices
                    }
                }
                g.tick(dt)
            }
            XCTAssertTrue(g.isOver)
            XCTAssertEqual(g.mistakes, 0, "skilled bot never errs (seed \(s))")
            grades.append(g.result().grade)
        }
        XCTAssertGreaterThanOrEqual(grades.filter { $0 >= .B }.count, seeds.count * 8 / 10, "\(grades)")

        var idle = ConvenienceGame(seed: 7)
        while !idle.isOver { idle.tick(dt) }
        XCTAssertLessThanOrEqual(idle.result().grade, .C)
        XCTAssertEqual(idle.result().payout, wage(.convenience, idle.result().grade))
    }

    func testConvenienceWrongChangeAndDoubleScan() {
        var g = ConvenienceGame(seed: 3)
        g.scan(0); g.scan(0)
        let afterDouble = g.mistakes
        XCTAssertGreaterThanOrEqual(afterDouble, 1)
        while g.phase == .scanning, let c = g.current, let i = c.scanned.firstIndex(of: false) { g.scan(i) }
        g.addChange(50_000); g.addChange(50_000)     // too much
        g.giveChange()
        XCTAssertEqual(g.mistakes, afterDouble + 1)
        XCTAssertEqual(g.changeDrawer, 0, "drawer is reset to count again")
    }

    // MARK: 카페
    func testCafeSkilledAndMixups() {
        var grades: [Grade] = []
        for s in seeds {
            var g = CafeGame(seed: s)
            var steps = 0
            while !g.isOver && steps < 10_000 {
                steps += 1
                if steps % 12 == 0, let o = g.orders.min(by: { $0.patience < $1.patience }) {   // ~1.2 s per drink
                    g.choose(o.spec.drink); g.setIced(o.spec.iced); g.setLarge(o.spec.large)
                    if g.cup.extraShot != o.spec.extraShot { g.toggleShot() }
                    g.serve(to: o.id)
                }
                g.tick(dt)
            }
            XCTAssertEqual(g.mixups, 0)
            grades.append(g.result().grade)
        }
        XCTAssertGreaterThanOrEqual(grades.filter { $0 >= .A }.count, seeds.count * 7 / 10, "\(grades)")

        var bad = CafeGame(seed: 2)
        while bad.orders.isEmpty { bad.tick(dt) }
        let o = bad.orders[0]
        bad.choose(o.spec.drink == .americano ? .latte : .americano)
        bad.serve(to: o.id)
        XCTAssertEqual(bad.mixups, 1)
        XCTAssertEqual(bad.orders.count, 1, "a wrong drink doesn't fulfil the order")
    }

    // MARK: 택배 상하차
    func testParcelGravityAndSkilledBot() {
        var t = ParcelGame.Truck(region: 0)
        XCTAssertEqual(t.restRow(col: 0, w: 2, h: 1), 0)
        t.place(col: 0, row: 0, w: 2, h: 1)
        XCTAssertEqual(t.restRow(col: 1, w: 2, h: 1), 1, "rests on top of the box below")
        XCTAssertNil(t.restRow(col: 4, w: 2, h: 1), "sticks out the side")
        XCTAssertEqual(t.restRow(col: 3, w: 2, h: 2), 0)

        var departures = 0, misSorted = 0, dropped = 0
        for s in seeds {
            var g = ParcelGame(seed: s)
            var steps = 0
            while !g.isOver && steps < 10_000 {
                steps += 1
                if steps % 10 == 0 {
                    let b = g.box, truck = g.trucks[b.region]
                    let cols = (0...(ParcelGame.width - b.w)).compactMap { c in truck.restRow(col: c, w: b.w, h: b.h).map { (c, $0) } }
                    if let best = cols.min(by: { $0.1 < $1.1 }) { g.place(truck: b.region, col: best.0) }
                }
                g.tick(dt)
            }
            departures += g.departures; misSorted += g.misSorted; dropped += g.dropped
        }
        XCTAssertEqual(misSorted, 0)
        XCTAssertGreaterThan(departures, seeds.count, "trucks fill and leave")
        XCTAssertLessThan(dropped, seeds.count * 3)
    }

    // MARK: 대리운전
    func testDriverSkilledArrivesAndIdlePaysForDamage() {
        var arrived = 0
        for s in seeds {
            var g = DriverGame(seed: s)
            while !g.isOver {
                // Pick the lane whose nearest obstacle ahead is farthest away, then steer toward its centre
                let lanes = [-0.66, 0, 0.66]
                func clearance(_ l: Double) -> Double { g.obstacles.filter { abs($0.x - l) < 0.2 && $0.y > -1 }.map { $0.y }.min() ?? 999 }
                let target = lanes.max(by: { clearance($0) < clearance($1) })!
                let err = target - g.x
                g.steer = abs(err) < 0.05 ? -g.vx * 0.5 : max(-1, min(1, err * 3 - g.vx))
                g.tick(dt)
            }
            if g.arrived { arrived += 1 }
            XCTAssertLessThanOrEqual(g.result().payout, wage(.designatedDriver, .S))
        }
        XCTAssertGreaterThanOrEqual(arrived, seeds.count * 7 / 10)

        var worst: Decimal = 0
        for s in seeds {
            var idle = DriverGame(seed: s)
            while !idle.isOver { idle.tick(dt) }
            let r = idle.result()
            XCTAssertEqual(r.payout, (idle.arrived ? wage(.designatedDriver, r.grade) : 0) - DriverGame.repairPerScrape * Decimal(idle.scrapes))
            worst = min(worst, r.payout)
        }
        XCTAssertLessThan(worst, 0, "scraping the customer's car costs you money")
        XCTAssertGreaterThanOrEqual(worst, -DriverGame.repairPerScrape * Decimal(DriverGame.maxScrapes))
    }

    // MARK: 서빙
    func testServingSkilledBalancesAndServes() {
        var served = 0, dropped = 0
        for s in seeds {
            var g = ServingGame(seed: s)
            var steps = 0
            while !g.isOver && steps < 10_000 {
                steps += 1
                if abs(g.tilt) > 0.25 { g.balance(g.tilt > 0 ? -1 : 1) }
                if steps % 8 == 0 {
                    if let t = g.tables.first(where: { tb in tb.wants != nil && g.tray.contains { $0.menu == tb.wants } }) {
                        g.serve(table: t.number)
                    } else if let d = g.pass.first(where: { d in g.tables.contains { $0.wants == d.menu } }), g.tray.count < 2 {
                        g.pickUp(d.id)
                    }
                }
                g.tick(dt)
            }
            served += g.served; dropped += g.droppedDishes
            XCTAssertEqual(g.wrongTable, 0)
        }
        XCTAssertGreaterThan(served, seeds.count * 6)
        XCTAssertLessThan(dropped, seeds.count)

        var clumsy = ServingGame(seed: 4)        // never balances, keeps a full tray
        var steps = 0
        while !clumsy.isOver && steps < 3000 {
            steps += 1
            if let d = clumsy.pass.first { clumsy.pickUp(d.id) }
            clumsy.tick(dt)
        }
        XCTAssertGreaterThan(clumsy.droppedDishes, 0, "an unbalanced full tray tips over")
    }

    // MARK: 전단지
    func testFlyersTimingAndBusyPeople() {
        var grades: [Grade] = []
        for s in seeds {
            var g = FlyerGame(seed: s)
            while !g.isOver {
                if let w = g.walkers.first(where: { !$0.done && $0.mood != .busy && abs($0.x - FlyerGame.zone) <= FlyerGame.perfectHalf }) { g.hand(to: w.id) }
                g.tick(dt / 2)
            }
            XCTAssertEqual(g.glares, 0)
            grades.append(g.result().grade)
        }
        XCTAssertGreaterThanOrEqual(grades.filter { $0 >= .B }.count, seeds.count * 7 / 10, "\(grades)")

        var spam = FlyerGame(seed: 9)            // hands constantly, to everyone
        while !spam.isOver { spam.hand(); spam.tick(dt) }
        XCTAssertGreaterThan(spam.ignored + spam.glares, 0)
    }

    // MARK: 과외
    func testTutoringQuestionsAreValidAndGraded() {
        for s in seeds {
            let g = TutoringGame(seed: s)
            XCTAssertEqual(g.questions.count, TutoringGame.count)
            for q in g.questions {
                XCTAssertEqual(Set(q.options).count, q.options.count, "no duplicate options: \(q)")
                XCTAssertTrue(q.options.indices.contains(q.answer))
                XCTAssertGreaterThanOrEqual(q.options.count, 3)
            }
        }
        var perfect = TutoringGame(seed: 5)
        while let q = perfect.current { perfect.answer(q.answer) }
        XCTAssertEqual(perfect.result().grade, .S)
        XCTAssertEqual(perfect.result().payout, wage(.tutoring, .S, extra: 20_000))

        var sleepy = TutoringGame(seed: 5)
        while !sleepy.isOver { sleepy.tick(0.5) }
        XCTAssertEqual(sleepy.correct, 0)
        XCTAssertEqual(sleepy.result().grade, .F)
    }

    func testTutoringMathAnswersAreCorrect() {
        var rng = GameRNG(seed: 11)
        for _ in 0..<200 {
            let q = TutoringGame.math(&rng, hard: true)
            let ans = Int(q.options[q.answer])!
            if q.prompt.contains("×") {
                let n = q.prompt.split(whereSeparator: { !$0.isNumber }).compactMap { Int($0) }
                XCTAssertEqual(n[0] * n[1], ans)
            } else {
                let n = q.prompt.split(whereSeparator: { !$0.isNumber }).compactMap { Int($0) }   // a, b, c
                XCTAssertEqual(n[0] * ans + n[1], n[2], q.prompt)
            }
        }
    }

    // MARK: 주식 단타
    func testDayTradingNeverLosesMoreThanStake() {
        var wins = 0, losses = 0
        for s in UInt64(1)...200 {
            let lev: Double = s % 3 == 0 ? 3 : 1
            var g = DayTradingGame(seed: s, stake: 5_000_000, leverage: lev)
            var rng = GameRNG(seed: s &* 31)
            while !g.isOver {
                if rng.chance(0.03) { g.holding ? g.sell() : g.buy() }
                g.tick(dt)
            }
            XCTAssertFalse(g.holding, "everything is sold at the bell")
            let r = g.result()
            XCTAssertGreaterThanOrEqual(r.payout, -5_000_000)
            if lev == 1 { XCTAssertGreaterThan(r.payout, -5_000_000 * 0.9, "1x can't be wiped out in a minute (±30% limits)") }
            if r.payout > 0 { wins += 1 } else if r.payout < 0 { losses += 1 }
        }
        XCTAssertGreaterThan(wins, 10); XCTAssertGreaterThan(losses, 10)    // genuinely risky both ways
    }

    func testDayTradingFeesAndDeterminism() {
        var a = DayTradingGame(seed: 42, stake: 1_000_000), b = DayTradingGame(seed: 42, stake: 1_000_000)
        for _ in 0..<50 { a.tick(dt); b.tick(dt) }
        XCTAssertEqual(a.price, b.price)
        a.buy(); a.sell()
        XCTAssertLessThan(a.equity, a.stakeValue, "a round trip at the same price costs fees + tax")
        XCTAssertEqual(DayTradingGame.stakes(for: 1), [300_000, 1_000_000])
        XCTAssertFalse(DayTradingGame.leverageAllowed(for: 2))
    }

    // MARK: Shared rules
    func testEligibilityWagesAndDailyLimit() {
        XCTAssertNil(JobKind.convenience.lockedReason(for: 0))
        XCTAssertNil(JobKind.dayTrading.lockedReason(for: 0))   // every job open to every tier
        XCTAssertNil(JobKind.tutoring.lockedReason(for: 0))
        XCTAssertNil(JobKind.tutoring.lockedReason(for: 2))
        XCTAssertNil(JobKind.parcel.lockedReason(for: 4))
        XCTAssertEqual(wage(.convenience, .B), 41_280)
        XCTAssertEqual(wage(.convenience, .S), 53_660)          // 41,280 × 1.3 = 53,664 → 53,660
        XCTAssertEqual(roundWon(12_345), 12_350)

        var st = JobsState()
        let r = ShiftResult(kind: .cafe, score: 500, grade: .A, payout: 48_300, lines: [])
        XCTAssertEqual(st.shiftsLeft(on: "2026-10-17"), JobsState.shiftsPerDay)
        st.record(r, day: "2026-10-17"); st.record(r, day: "2026-10-17")
        XCTAssertEqual(st.shiftsLeft(on: "2026-10-17"), JobsState.shiftsPerDay - 2)
        XCTAssertEqual(st.shiftsLeft(on: "2026-10-18"), JobsState.shiftsPerDay, "a new day restores energy")
        XCTAssertEqual(st.best["cafe"], 500)
        XCTAssertEqual(st.earned["cafe"], 96_600)
        let json = try! JSONEncoder().encode(st)
        XCTAssertEqual(try! JSONDecoder().decode(JobsState.self, from: json), st)
    }
}

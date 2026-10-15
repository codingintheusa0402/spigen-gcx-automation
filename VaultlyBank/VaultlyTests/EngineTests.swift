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
        XCTAssertEqual(Money.format(1_234_567, code: "KRW"), "₩1,234,567")
        XCTAssertTrue(Money.format(5000, code: "KRW", signed: true).hasPrefix("+"))
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

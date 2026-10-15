import Foundation
import Combine
import SwiftUI

struct BannerMessage: Identifiable, Equatable {
    let id = UUID()
    let title: String
    let body: String
    let positive: Bool
}

/// Single source of truth. Every mutation goes through here so it is persisted
/// and the notification schedule stays in sync.
@MainActor
final class BankStore: ObservableObject {
    @Published private(set) var data: BankData {
        didSet { if data != oldValue { Persistence.save(data) } }
    }
    @Published var banner: BannerMessage?
    @Published var notificationsAuthorized = false

    private var timer: AnyCancellable?

    init() {
        data = Persistence.load() ?? BankData.starter()
        // Check for due events every 20s while the app is in the foreground so the
        // "4:00pm allowance" shows up on screen within seconds of the notification.
        timer = Timer.publish(every: 20, on: .main, in: .common).autoconnect()
            .sink { [weak self] _ in Task { @MainActor in self?.catchUp() } }
        Task { notificationsAuthorized = await NotificationService.authorizationStatus() == .authorized }
    }

    // MARK: Derived

    var settings: BankSettings { data.settings }
    var currency: String { data.settings.currencyCode }
    var theme: Theme { Theme.named(data.settings.themeID) }
    var checkingAccounts: [Account] { data.accounts.filter { $0.kind == .checking } }
    var productAccounts: [Account] { data.accounts.filter { $0.kind.isProduct } }

    var bankAssets: Decimal {
        data.accounts.reduce(0) { $0 + ($1.kind == .loan ? -$1.balance : $1.balance) }
    }
    var totalAssets: Decimal { bankAssets + data.portfolio.investedKRW + data.portfolio.cashKRWTotal }
    var portfolio: Portfolio { data.portfolio }
    var totalDebt: Decimal { data.accounts.filter { $0.kind == .loan }.reduce(0) { $0 + $1.balance } }

    func account(_ id: UUID) -> Account? { data.accounts.first { $0.id == id } }

    func transactions(for id: UUID) -> [Transaction] {
        data.transactions.filter { $0.accountID == id }.sorted { $0.date > $1.date }
    }
    var recentTransactions: [Transaction] { Array(data.transactions.sorted { $0.date > $1.date }.prefix(30)) }

    // MARK: Time-based settlement

    func catchUp() {
        var d = data
        let fired = ScheduleEngine.processDue(now: Date(), data: &d)
        InterestEngine.accrue(now: Date(), data: &d, taxRate: d.settings.interestTaxRate)
        MarketSimulator.tick(now: Date(), portfolio: &d.portfolio)
        LifestyleEngine.process(now: Date(), data: &d)
        if d != data { data = d }
        if let last = fired.last {
            let amount = Money.format(abs(last.event.amount), code: currency)
            banner = BannerMessage(title: last.event.isIncoming ? "입금 \(amount)" : "출금 \(amount)",
                                   body: "\(last.event.counterparty) · \(last.event.memo)",
                                   positive: last.event.isIncoming)
        }
        syncNotifications()
    }

    func syncNotifications() {
        NotificationService.sync(events: data.events, accounts: data.accounts, currency: currency)
    }

    func requestNotifications() {
        Task {
            notificationsAuthorized = await NotificationService.requestAuthorization()
            syncNotifications()
        }
    }

    // MARK: Settings

    func update(_ change: (inout BankSettings) -> Void) {
        var s = data.settings; change(&s); data.settings = s
    }

    func resetAll() {
        Persistence.reset()
        data = BankData.starter()
        syncNotifications()
    }

    // MARK: Accounts

    func addChecking(name: String, balance: Decimal, colorIndex: Int) {
        var d = data
        let acc = Account(name: name, kind: .checking, number: Account.generateNumber(), balance: 0, colorIndex: colorIndex)
        d.accounts.append(acc)
        if balance != 0 {
            Ledger.apply(amount: balance, to: acc.id, date: Date(), counterparty: "직접 설정",
                         memo: "시작 잔액", category: .adjustment, data: &d)
        }
        data = d
    }

    func setBalance(_ balance: Decimal, of id: UUID) {
        var d = data; Ledger.setBalance(balance, of: id, data: &d); data = d
    }

    func rename(_ id: UUID, to name: String) {
        guard let i = data.accounts.firstIndex(where: { $0.id == id }) else { return }
        data.accounts[i].name = name
    }

    func deleteAccount(_ id: UUID) {
        var d = data
        d.accounts.removeAll { $0.id == id }
        d.transactions.removeAll { $0.accountID == id }
        d.events.removeAll { $0.accountID == id }
        if d.lifestyle.accountID == id { d.lifestyle.accountID = nil; d.lifestyle.tier = nil }
        for i in d.accounts.indices where d.accounts[i].terms?.linkedAccountID == id {
            d.accounts[i].terms?.linkedAccountID = nil
        }
        data = d
        syncNotifications()
    }

    // MARK: Transfers (mock)

    enum TransferTarget { case own(UUID), external(name: String, bank: String) }

    func transfer(from: UUID, to target: TransferTarget, amount: Decimal, memo: String) -> Bool {
        guard amount > 0, let src = account(from), src.balance >= amount else { return false }
        var d = data
        switch target {
        case .own(let dst):
            guard let dstAcc = account(dst), dst != from else { return false }
            Ledger.apply(amount: -amount, to: from, date: Date(), counterparty: dstAcc.name,
                         memo: memo.isEmpty ? "내 계좌 이체" : memo, category: .transfer, data: &d)
            Ledger.apply(amount: amount, to: dst, date: Date(), counterparty: src.name,
                         memo: memo.isEmpty ? "내 계좌 이체" : memo, category: .transfer, data: &d)
        case .external(let name, let bank):
            Ledger.apply(amount: -amount, to: from, date: Date(), counterparty: name,
                         memo: memo.isEmpty ? bank : "\(bank) · \(memo)", category: .transfer, data: &d)
        }
        data = d
        return true
    }

    // MARK: Scheduled events

    func addEvent(_ ev: ScheduledEvent) { data.events.append(ev); syncNotifications() }
    func updateEvent(_ ev: ScheduledEvent) {
        guard let i = data.events.firstIndex(where: { $0.id == ev.id }) else { return }
        data.events[i] = ev; syncNotifications()
    }
    func deleteEvent(_ id: UUID) { data.events.removeAll { $0.id == id }; syncNotifications() }
    func toggleEvent(_ id: UUID) {
        guard let i = data.events.firstIndex(where: { $0.id == id }) else { return }
        data.events[i].isEnabled.toggle()
        // re-enabling a past one-off: push it to the same time tomorrow
        if data.events[i].isEnabled, data.events[i].nextFireDate <= Date(), data.events[i].repeatRule == .none,
           let t = Calendar.current.date(byAdding: .day, value: 1, to: data.events[i].nextFireDate) {
            data.events[i].nextFireDate = t
        }
        syncNotifications()
    }

    // MARK: Lifestyle (wealth tier)

    var lifestyle: LifestyleState { data.lifestyle }

    /// Pick a wealth tier. Optionally resets the account balance to the tier's
    /// typical amount and back-fills the last `backfillDays` of transactions.
    func applyTier(_ tier: WealthTier?, accountID: UUID?, resetBalance: Bool, backfillDays: Int) {
        var d = data
        d.lifestyle.tier = tier
        d.lifestyle.accountID = accountID
        d.lifestyle.seed = UInt64.random(in: 1...UInt64.max)
        if let tier, let accountID {
            if resetBalance { Ledger.setBalance(tier.startingBalance, of: accountID, data: &d) }
            d.lifestyle.lastProcessed = Calendar.current.date(byAdding: .day, value: -backfillDays, to: Date()) ?? Date()
            if backfillDays > 0 {
                // Re-date the balance reset so it precedes the generated history
                if resetBalance, let i = d.transactions.lastIndex(where: { $0.accountID == accountID && $0.category == .adjustment }) {
                    d.transactions[i].date = d.lifestyle.lastProcessed
                }
                LifestyleEngine.process(now: Date(), data: &d)
            }
        } else {
            d.lifestyle.lastProcessed = Date()
        }
        data = d
    }

    // MARK: Mock trading

    func updatePortfolio(_ change: (inout Portfolio) -> Void) {
        var p = data.portfolio; change(&p); data.portfolio = p
    }

    func addHolding(_ h: Holding) { updatePortfolio { $0.holdings.append(h) } }

    func removeHolding(_ id: UUID) {
        updatePortfolio { $0.holdings.removeAll { $0.id == id }; $0.trades.removeAll { $0.holdingID == id } }
    }

    func setPrice(_ price: Decimal, of id: UUID) {
        updatePortfolio { p in
            if let i = p.holdings.firstIndex(where: { $0.id == id }) { p.holdings[i].price = price }
        }
    }

    /// Dry-run order. Buys draw from mock cash in the holding's currency; sells add to it.
    func trade(holdingID: UUID, buy: Bool, shares: Decimal, price: Decimal) -> Bool {
        guard shares > 0, price > 0 else { return false }
        var p = data.portfolio
        guard let i = p.holdings.firstIndex(where: { $0.id == holdingID }) else { return false }
        let total = (shares * price).rounded(scale: 2)
        var h = p.holdings[i]
        if buy {
            if h.market == .domestic { guard p.cashKRW >= total else { return false }; p.cashKRW -= total }
            else { guard p.cashUSD >= total else { return false }; p.cashUSD -= total }
            h.avgCost = ((h.cost + total) / (h.shares + shares)).rounded(scale: 2)
            h.shares += shares
        } else {
            guard h.shares >= shares else { return false }
            if h.market == .domestic { p.cashKRW += total } else { p.cashUSD += total }
            h.shares -= shares
        }
        h.price = price
        p.holdings[i] = h
        p.trades.append(TradeRecord(holdingID: holdingID, date: Date(), isBuy: buy, shares: shares, price: price))
        data.portfolio = p
        return true
    }

    // MARK: Products

    func openProduct(kind: AccountKind, name: String, annualRate: Double, termMonths: Int,
                     amount: Decimal, linkedAccountID: UUID?, startDate: Date) {
        var d = data
        let number = Account.generateNumber()
        var terms = ProductTerms(annualRate: annualRate, termMonths: termMonths, startDate: startDate,
                                 principal: 0, monthlyAmount: 0, linkedAccountID: linkedAccountID,
                                 lastAccrualDate: startDate, lastInstallmentDate: startDate)
        var acc = Account(name: name, kind: kind, number: number, balance: 0, terms: nil, colorIndex: Int.random(in: 0..<6))
        switch kind {
        case .deposit:
            terms.principal = amount
            d.accounts.append(acc)
            if let linked = linkedAccountID { // move the lump sum out of checking
                Ledger.apply(amount: -amount, to: linked, date: startDate, counterparty: name, memo: "예금 가입", category: .transfer, data: &d)
            }
            Ledger.apply(amount: amount, to: acc.id, date: startDate, counterparty: "가입", memo: "예치금", category: .transfer, data: &d)
        case .installmentSavings:
            terms.monthlyAmount = amount
            d.accounts.append(acc)
            // first installment on the start date
            if let linked = linkedAccountID {
                Ledger.apply(amount: -amount, to: linked, date: startDate, counterparty: name, memo: "적금 자동이체", category: .autoTransfer, data: &d)
            }
            Ledger.apply(amount: amount, to: acc.id, date: startDate, counterparty: "1회차", memo: "월 납입", category: .autoTransfer, data: &d)
        case .loan:
            terms.principal = amount
            terms.monthlyAmount = InterestEngine.annuityPayment(principal: amount, annualRate: annualRate, months: termMonths)
            acc.balance = 0
            d.accounts.append(acc)
            Ledger.apply(amount: amount, to: acc.id, date: startDate, counterparty: "대출 실행", memo: "원금", category: .transfer, data: &d)
            if let linked = linkedAccountID {
                Ledger.apply(amount: amount, to: linked, date: startDate, counterparty: name, memo: "대출금 입금", category: .transfer, data: &d)
            }
        case .checking:
            return
        }
        if let i = d.accounts.firstIndex(where: { $0.id == acc.id }) { d.accounts[i].terms = terms }
        data = d
        catchUp() // if startDate is in the past, settle up immediately
    }
}

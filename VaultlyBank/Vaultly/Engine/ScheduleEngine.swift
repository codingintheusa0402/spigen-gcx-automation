import Foundation

/// Applies due scheduled events to the ledger. Runs on launch / foreground / timer so
/// events still land even if the app was closed when the time passed.
enum ScheduleEngine {
    struct Fired { let event: ScheduledEvent; let transaction: Transaction }

    @discardableResult
    static func processDue(now: Date = Date(), data: inout BankData) -> [Fired] {
        var fired: [Fired] = []
        for i in data.events.indices {
            var ev = data.events[i]
            guard ev.isEnabled else { continue }
            var guardCount = 0
            while ev.nextFireDate <= now, guardCount < 5000 {
                guardCount += 1
                if let to = ev.toAccountID {
                    // internal transfer (savings goal auto-transfer): out of accountID, into toAccountID
                    let amount = abs(ev.amount)
                    let toName = data.accounts.first { $0.id == to }?.name ?? ev.counterparty
                    let fromName = data.accounts.first { $0.id == ev.accountID }?.name ?? "내 계좌"
                    if let tx = Ledger.apply(amount: -amount, to: ev.accountID, date: ev.nextFireDate, counterparty: toName,
                                             memo: ev.memo, category: .autoTransfer, tag: SpendingCategory.transfer.rawValue, data: &data) {
                        Ledger.apply(amount: amount, to: to, date: ev.nextFireDate, counterparty: fromName,
                                     memo: ev.memo, category: .autoTransfer, tag: SpendingCategory.transfer.rawValue, data: &data)
                        fired.append(Fired(event: ev, transaction: tx))
                    }
                } else if let tx = Ledger.apply(amount: ev.amount, to: ev.accountID, date: ev.nextFireDate,
                                                counterparty: ev.counterparty, memo: ev.memo,
                                                category: .scheduled, data: &data) {
                    fired.append(Fired(event: ev, transaction: tx))
                }
                ev.lastFiredAt = ev.nextFireDate
                if let next = ev.repeatRule.next(after: ev.nextFireDate) {
                    ev.nextFireDate = next
                } else {
                    ev.isEnabled = false
                    break
                }
            }
            data.events[i] = ev
        }
        return fired
    }
}

/// Single place that mutates balances so every change leaves a transaction behind.
enum Ledger {
    @discardableResult
    static func apply(amount: Decimal, to accountID: UUID, date: Date, counterparty: String,
                      memo: String, category: TransactionCategory, tag: String? = nil, data: inout BankData) -> Transaction? {
        guard let idx = data.accounts.firstIndex(where: { $0.id == accountID }) else { return nil }
        data.accounts[idx].balance += amount
        // App Store build: real brand names become fictional ones here, the one place every transaction is made
        let tx = Transaction(accountID: accountID, date: date, amount: amount, counterparty: StoreNames.fix(counterparty),
                             memo: StoreNames.fix(memo), category: category, balanceAfter: data.accounts[idx].balance, tag: tag)
        data.transactions.append(tx)
        return tx
    }

    static func setBalance(_ newBalance: Decimal, of accountID: UUID, date: Date = Date(), data: inout BankData) {
        guard let idx = data.accounts.firstIndex(where: { $0.id == accountID }) else { return }
        let diff = newBalance - data.accounts[idx].balance
        guard diff != 0 else { return }
        apply(amount: diff, to: accountID, date: date, counterparty: "직접 설정",
              memo: "잔액 조정", category: .adjustment, data: &data)
    }
}

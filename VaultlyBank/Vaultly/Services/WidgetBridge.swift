import Foundation
import WidgetKit

/// Publishes a small snapshot to the app group so the widgets can draw without the whole bank document.
enum WidgetBridge {
    @MainActor
    static func publish(store: BankStore) {
        let main = store.checkingAccounts.first
        let now = store.now
        var snap = WidgetSnapshot()
        snap.bank = store.brand.bank
        snap.accountName = main?.name ?? "입출금통장"
        snap.balance = main?.balance ?? 0
        snap.totalAssets = store.totalAssets
        snap.wonSuffix = store.theme.wonSuffix
        snap.currency = store.currency
        snap.hideBalances = store.settings.hideBalances
        snap.streak = store.challenge.streak(now: now)
        snap.monthSaved = store.challenge.total(inMonthOf: now)
        snap.monthSpent = store.report(for: now).spent
        if let ev = store.nextIncomingEvent {
            snap.nextCounterparty = ev.counterparty
            snap.nextAmount = ev.amount
            snap.nextDate = ev.nextFireDate
        }
        snap.updatedAt = Date()
        snap.save()
        WidgetCenter.shared.reloadAllTimelines()
    }
}

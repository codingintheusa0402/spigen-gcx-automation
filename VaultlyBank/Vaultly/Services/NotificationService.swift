import Foundation
import UserNotifications

/// Local notifications for scheduled events. One pending request per enabled event;
/// repeating events use calendar triggers so they keep firing without the app running.
enum NotificationService {
    private static let prefix = "vaultly.event."

    static func requestAuthorization() async -> Bool {
        let center = UNUserNotificationCenter.current()
        return (try? await center.requestAuthorization(options: [.alert, .sound, .badge])) ?? false
    }

    static func authorizationStatus() async -> UNAuthorizationStatus {
        await UNUserNotificationCenter.current().notificationSettings().authorizationStatus
    }

    /// "[우리은행] 03/09 15:40 입금 31,440,000원 엄마 잔액 123,760,684,504원" — the line real Korean bank pushes use.
    static func bankLine(bank: String, date: Date, incoming: Bool, amount: Decimal, counterparty: String, balance: Decimal, currency: String) -> String {
        let f = DateFormatter(); f.dateFormat = "MM/dd HH:mm"
        let bal = Money.format(balance, code: currency)
        return "[\(bank)] \(f.string(from: date)) \(incoming ? "입금" : "출금") \(Money.format(amount, code: currency)) \(counterparty) 잔액 \(bal)"
    }

    static func sync(events: [ScheduledEvent], accounts: [Account], bank: String, currency: String) {
        let center = UNUserNotificationCenter.current()
        center.getPendingNotificationRequests { pending in
            let stale = pending.map(\.identifier).filter { $0.hasPrefix(prefix) }
            center.removePendingNotificationRequests(withIdentifiers: stale)
            for ev in events where ev.isEnabled && ev.nextFireDate > Date() {
                let content = UNMutableNotificationContent()
                let account = accounts.first { $0.id == ev.accountID }
                let incoming = ev.isIncoming && !ev.isInternal
                let projected = (account?.balance ?? 0) + (incoming ? abs(ev.amount) : -abs(ev.amount))
                content.title = ev.isInternal ? "자동이체" : (incoming ? "입금" : "출금")
                content.body = bankLine(bank: bank, date: ev.nextFireDate, incoming: incoming, amount: abs(ev.amount),
                                        counterparty: ev.counterparty, balance: projected, currency: currency)
                content.sound = .default
                content.interruptionLevel = .timeSensitive

                var comps: DateComponents
                let cal = Calendar.current
                switch ev.repeatRule {
                case .none:
                    comps = cal.dateComponents([.year, .month, .day, .hour, .minute], from: ev.nextFireDate)
                case .daily:
                    comps = cal.dateComponents([.hour, .minute], from: ev.nextFireDate)
                case .weekly:
                    comps = cal.dateComponents([.weekday, .hour, .minute], from: ev.nextFireDate)
                case .monthly:
                    comps = cal.dateComponents([.day, .hour, .minute], from: ev.nextFireDate)
                }
                let trigger = UNCalendarNotificationTrigger(dateMatching: comps, repeats: ev.repeatRule != .none)
                let req = UNNotificationRequest(identifier: prefix + ev.id.uuidString, content: content, trigger: trigger)
                center.add(req)
            }
        }
    }
}

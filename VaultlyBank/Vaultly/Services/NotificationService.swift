import Foundation
import UserNotifications

/// Local notifications for scheduled events. One pending request per enabled event;
/// repeating events use calendar triggers so they keep firing without the app running.
enum NotificationService {
    private static let prefix = "vaultly.event."

    static let questCategory = "VAULTLY_QUEST"
    static let questPrefix = "vaultly.quest."

    static func requestAuthorization() async -> Bool {
        let center = UNUserNotificationCenter.current()
        registerCategories()
        return (try? await center.requestAuthorization(options: [.alert, .sound, .badge])) ?? false
    }

    static func registerCategories() {
        let actions = (0..<3).map { UNNotificationAction(identifier: "opt\($0)", title: "선택 \($0 + 1)", options: [.foreground]) }
        let cat = UNNotificationCategory(identifier: questCategory, actions: actions, intentIdentifiers: [], options: [])
        UNUserNotificationCenter.current().setNotificationCategories([cat])
    }

    /// "새로운 선택" alert outside the app. Options are listed in the body and offered as action buttons.
    static func syncQuest(_ upcoming: (Quest, Date)?, pendingAlreadyShown: Bool) {
        let center = UNUserNotificationCenter.current()
        center.getPendingNotificationRequests { pending in
            center.removePendingNotificationRequests(withIdentifiers: pending.map(\.identifier).filter { $0.hasPrefix(questPrefix) })
            guard let (q, fire) = upcoming, fire.timeIntervalSinceNow > (pendingAlreadyShown ? 0 : 20) else { return }
            let content = UNMutableNotificationContent()
            content.title = "새로운 선택 · \(q.category.title)"
            content.body = q.title + "\n" + q.options.enumerated().map { "\($0.offset + 1). \($0.element.title)" }.joined(separator: "  ")
            content.sound = .default
            content.categoryIdentifier = questCategory
            content.userInfo = ["questID": q.id]
            content.interruptionLevel = .timeSensitive
            let trigger = UNTimeIntervalNotificationTrigger(timeInterval: max(1, fire.timeIntervalSinceNow), repeats: false)
            center.add(UNNotificationRequest(identifier: questPrefix + q.id, content: content, trigger: trigger))
        }
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

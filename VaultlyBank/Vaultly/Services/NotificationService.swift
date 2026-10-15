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

    static func sync(events: [ScheduledEvent], accounts: [Account], currency: String) {
        let center = UNUserNotificationCenter.current()
        center.getPendingNotificationRequests { pending in
            let stale = pending.map(\.identifier).filter { $0.hasPrefix(prefix) }
            center.removePendingNotificationRequests(withIdentifiers: stale)
            for ev in events where ev.isEnabled && ev.nextFireDate > Date() {
                let content = UNMutableNotificationContent()
                let account = accounts.first { $0.id == ev.accountID }
                let amount = Money.format(abs(ev.amount), code: currency)
                content.title = ev.isIncoming ? "입금 \(amount)" : "출금 \(amount)"
                content.body = ev.isIncoming
                    ? "\(ev.counterparty)님이 \(amount)을 보냈습니다 · \(ev.memo)"
                    : "\(ev.counterparty) · \(ev.memo) \(amount)"
                content.subtitle = account?.name ?? ""
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

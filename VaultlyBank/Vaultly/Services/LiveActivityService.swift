import Foundation
import ActivityKit

/// Shows a Live Activity for the next incoming deposit when it is a few hours away,
/// so "입금 예정 16:00" sits on the lock screen / Dynamic Island until it lands.
enum LiveActivityService {
    static let horizon: TimeInterval = 4 * 3600

    static func sync(events: [ScheduledEvent], accounts: [Account], bank: String, currency: String) {
        guard ActivityAuthorizationInfo().areActivitiesEnabled else { return }
        let now = Date()
        let next = events
            .filter { $0.isEnabled && $0.isIncoming && !$0.isInternal && $0.nextFireDate > now && $0.nextFireDate.timeIntervalSince(now) <= horizon }
            .min { $0.nextFireDate < $1.nextFireDate }
        Task {
            for activity in Activity<DepositActivityAttributes>.activities {
                if let next, activity.attributes.eventID == next.id.uuidString,
                   activity.content.state.fireDate == next.nextFireDate { continue }
                // landed or no longer relevant
                let landed = activity.content.state.fireDate <= now
                let state = DepositActivityAttributes.ContentState(fireDate: activity.content.state.fireDate, landed: landed)
                await activity.end(ActivityContent(state: state, staleDate: nil), dismissalPolicy: landed ? .after(now.addingTimeInterval(1800)) : .immediate)
            }
            guard let next,
                  !Activity<DepositActivityAttributes>.activities.contains(where: { $0.attributes.eventID == next.id.uuidString && $0.content.state.fireDate == next.nextFireDate })
            else { return }
            let account = accounts.first { $0.id == next.accountID }
            let attrs = DepositActivityAttributes(bank: bank, counterparty: next.counterparty, amountText: Money.format(abs(next.amount), code: currency),
                                                  accountName: account?.name ?? "", eventID: next.id.uuidString)
            let content = ActivityContent(state: DepositActivityAttributes.ContentState(fireDate: next.nextFireDate, landed: false), staleDate: next.nextFireDate.addingTimeInterval(600))
            _ = try? Activity.request(attributes: attrs, content: content)
        }
    }
}

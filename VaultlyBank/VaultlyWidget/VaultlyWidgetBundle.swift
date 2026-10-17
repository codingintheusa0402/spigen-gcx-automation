import WidgetKit
import SwiftUI

@main
struct VaultlyWidgetBundle: WidgetBundle {
    var body: some Widget {
        BalanceWidget()
        ChallengeWidget()
        DepositLiveActivity()
    }
}

struct SnapshotEntry: TimelineEntry {
    let date: Date
    let snap: WidgetSnapshot
}

struct SnapshotProvider: TimelineProvider {
    func placeholder(in context: Context) -> SnapshotEntry {
        var s = WidgetSnapshot(); s.balance = 1_250_000; s.totalAssets = 4_370_000; s.streak = 7; s.monthSaved = 184_000; s.monthSpent = 1_230_000
        return SnapshotEntry(date: Date(), snap: s)
    }
    func getSnapshot(in context: Context, completion: @escaping (SnapshotEntry) -> Void) {
        completion(SnapshotEntry(date: Date(), snap: WidgetSnapshot.load() ?? placeholder(in: context).snap))
    }
    func getTimeline(in context: Context, completion: @escaping (Timeline<SnapshotEntry>) -> Void) {
        let entry = SnapshotEntry(date: Date(), snap: WidgetSnapshot.load() ?? placeholder(in: context).snap)
        completion(Timeline(entries: [entry], policy: .after(Date().addingTimeInterval(1800))))
    }
}

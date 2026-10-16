import WidgetKit
import SwiftUI

struct BalanceWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: WidgetShared.kind, provider: SnapshotProvider()) { entry in
            BalanceWidgetView(snap: entry.snap)
                .containerBackground(for: .widget) { Color(.systemBackground) }
        }
        .configurationDisplayName("잔액")
        .description("입출금통장 잔액과 다음 입금 예정을 보여줍니다.")
        .supportedFamilies([.systemSmall, .systemMedium, .accessoryRectangular, .accessoryCircular, .accessoryInline])
    }
}

struct BalanceWidgetView: View {
    @Environment(\.widgetFamily) private var family
    let snap: WidgetSnapshot

    var body: some View {
        switch family {
        case .accessoryInline:
            Text("잔액 \(snap.money(snap.balance))")
        case .accessoryCircular:
            ZStack {
                AccessoryWidgetBackground()
                VStack(spacing: 0) {
                    Image(systemName: "building.columns.fill").font(.system(size: 12, weight: .semibold))
                    Text(short(snap.balance)).font(.custom("Pretendard-SemiBold", size: 12)).minimumScaleFactor(0.6).lineLimit(1)
                }
            }
        case .accessoryRectangular:
            VStack(alignment: .leading, spacing: 2) {
                Text(snap.accountName).font(.custom("Pretendard-Medium", size: 12)).foregroundStyle(.secondary)
                Text(snap.money(snap.balance)).font(.custom("Pretendard-SemiBold", size: 16)).minimumScaleFactor(0.6).lineLimit(1)
                if let cp = snap.nextCounterparty, let a = snap.nextAmount, let d = snap.nextDate, d > Date() {
                    Text("\(d.formatted(date: .omitted, time: .shortened)) \(cp) \(snap.money(a))").font(.custom("Pretendard-Regular", size: 11)).foregroundStyle(.secondary).lineLimit(1)
                }
            }
        case .systemMedium:
            HStack(alignment: .top, spacing: 14) {
                smallBody
                Divider()
                VStack(alignment: .leading, spacing: 8) {
                    labeled("총자산", snap.money(snap.totalAssets))
                    labeled("이번 달 지출", snap.money(snap.monthSpent))
                    if let cp = snap.nextCounterparty, let a = snap.nextAmount, let d = snap.nextDate, d > Date() {
                        VStack(alignment: .leading, spacing: 2) {
                            Text("다음 입금").font(.custom("Pretendard-Regular", size: 11)).foregroundStyle(.secondary)
                            Text("\(d.formatted(date: .omitted, time: .shortened)) \(cp)").font(.custom("Pretendard-SemiBold", size: 13)).lineLimit(1)
                            Text(snap.money(a)).font(.custom("Pretendard-SemiBold", size: 13)).foregroundStyle(.blue)
                        }
                    }
                    Spacer(minLength: 0)
                }
            }
        default:
            smallBody
        }
    }

    private var smallBody: some View {
        VStack(alignment: .leading, spacing: 6) {
            HStack(spacing: 6) {
                Image(systemName: "building.columns.fill").font(.system(size: 12, weight: .semibold)).foregroundStyle(.blue)
                Text(snap.bank).font(.custom("Pretendard-SemiBold", size: 12)).foregroundStyle(.secondary).lineLimit(1)
            }
            Spacer(minLength: 0)
            Text(snap.accountName).font(.custom("Pretendard-Regular", size: 12)).foregroundStyle(.secondary).lineLimit(1)
            Text(snap.money(snap.balance)).font(.custom("Pretendard-SemiBold", size: 22)).minimumScaleFactor(0.5).lineLimit(1)
        }
        .frame(maxWidth: .infinity, maxHeight: .infinity, alignment: .leading)
    }

    private func labeled(_ t: String, _ v: String) -> some View {
        VStack(alignment: .leading, spacing: 2) {
            Text(t).font(.custom("Pretendard-Regular", size: 11)).foregroundStyle(.secondary)
            Text(v).font(.custom("Pretendard-SemiBold", size: 14)).lineLimit(1).minimumScaleFactor(0.6)
        }
    }

    private func short(_ d: Decimal) -> String {
        if snap.hideBalances { return "••" }
        let v = NSDecimalNumber(decimal: d).doubleValue
        if v >= 100_000_000 { return String(format: "%.1f억", v / 100_000_000) }
        if v >= 10_000 { return String(format: "%.0f만", v / 10_000) }
        return String(format: "%.0f", v)
    }
}

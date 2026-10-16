import WidgetKit
import SwiftUI

struct ChallengeWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: WidgetShared.challengeKind, provider: SnapshotProvider()) { entry in
            ChallengeWidgetView(snap: entry.snap)
                .containerBackground(for: .widget) { Color(.systemBackground) }
        }
        .configurationDisplayName("무지출 챌린지")
        .description("연속 일수와 이번 달 참은 돈.")
        .supportedFamilies([.systemSmall, .accessoryRectangular, .accessoryCircular])
    }
}

struct ChallengeWidgetView: View {
    @Environment(\.widgetFamily) private var family
    let snap: WidgetSnapshot

    var body: some View {
        switch family {
        case .accessoryCircular:
            ZStack {
                AccessoryWidgetBackground()
                VStack(spacing: 0) {
                    Image(systemName: "flame.fill").font(.system(size: 13, weight: .semibold))
                    Text("\(snap.streak)일").font(.custom("Pretendard-SemiBold", size: 12))
                }
            }
        case .accessoryRectangular:
            VStack(alignment: .leading, spacing: 2) {
                Label("무지출 \(snap.streak)일 연속", systemImage: "flame.fill").font(.custom("Pretendard-SemiBold", size: 13))
                Text("이번 달 참은 돈").font(.custom("Pretendard-Regular", size: 11)).foregroundStyle(.secondary)
                Text(snap.money(snap.monthSaved)).font(.custom("Pretendard-SemiBold", size: 15)).lineLimit(1).minimumScaleFactor(0.6)
            }
        default:
            VStack(alignment: .leading, spacing: 6) {
                HStack(spacing: 6) {
                    Image(systemName: "flame.fill").foregroundStyle(.orange)
                    Text("무지출 챌린지").font(.custom("Pretendard-SemiBold", size: 12)).foregroundStyle(.secondary)
                }
                Spacer(minLength: 0)
                Text("\(snap.streak)일 연속").font(.custom("Pretendard-Bold", size: 20))
                Text("이번 달 \(snap.money(snap.monthSaved))").font(.custom("Pretendard-Medium", size: 12)).foregroundStyle(.secondary).lineLimit(1).minimumScaleFactor(0.6)
            }
            .frame(maxWidth: .infinity, maxHeight: .infinity, alignment: .leading)
        }
    }
}

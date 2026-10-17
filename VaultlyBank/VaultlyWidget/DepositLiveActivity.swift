import ActivityKit
import WidgetKit
import SwiftUI

/// Lock screen / Dynamic Island: "입금 예정 16:00 · 엄마 5,000,000원" counting down, then "입금 완료".
struct DepositLiveActivity: Widget {
    var body: some WidgetConfiguration {
        ActivityConfiguration(for: DepositActivityAttributes.self) { context in
            lockScreen(context)
                .activityBackgroundTint(Color(.systemBackground).opacity(0.85))
        } dynamicIsland: { context in
            DynamicIsland {
                DynamicIslandExpandedRegion(.leading) {
                    Label(context.attributes.counterparty, systemImage: "arrow.down.left.circle.fill").font(.custom("Pretendard-SemiBold", size: 14))
                }
                DynamicIslandExpandedRegion(.trailing) {
                    Text(context.attributes.amountText).font(.custom("Pretendard-Bold", size: 16)).foregroundStyle(.blue)
                }
                DynamicIslandExpandedRegion(.bottom) {
                    if context.state.landed {
                        Text("[\(context.attributes.bank)] 입금 완료 · \(context.attributes.accountName)").font(.custom("Pretendard-Regular", size: 12)).foregroundStyle(.secondary)
                    } else {
                        HStack {
                            Text("[\(context.attributes.bank)] 입금 예정 \(context.state.fireDate.formatted(date: .omitted, time: .shortened))")
                            Spacer()
                            Text(timerInterval: Date()...context.state.fireDate, countsDown: true).monospacedDigit()
                        }
                        .font(.custom("Pretendard-Regular", size: 12)).foregroundStyle(.secondary)
                    }
                }
            } compactLeading: {
                Image(systemName: context.state.landed ? "checkmark.circle.fill" : "arrow.down.circle").foregroundStyle(.blue)
            } compactTrailing: {
                if context.state.landed { Text("입금").font(.custom("Pretendard-SemiBold", size: 12)) }
                else { Text(timerInterval: Date()...context.state.fireDate, countsDown: true).monospacedDigit().font(.custom("Pretendard-SemiBold", size: 12)).frame(width: 48) }
            } minimal: {
                Image(systemName: "arrow.down.circle.fill").foregroundStyle(.blue)
            }
        }
    }

    private func lockScreen(_ context: ActivityViewContext<DepositActivityAttributes>) -> some View {
        HStack(spacing: 14) {
            ZStack {
                Circle().fill(Color.blue.opacity(0.15)).frame(width: 44, height: 44)
                Image(systemName: context.state.landed ? "checkmark" : "arrow.down").font(.system(size: 18, weight: .bold)).foregroundStyle(.blue)
            }
            VStack(alignment: .leading, spacing: 3) {
                Text(context.state.landed ? "입금 완료" : "입금 예정 \(context.state.fireDate.formatted(date: .omitted, time: .shortened))")
                    .font(.custom("Pretendard-SemiBold", size: 15))
                Text("[\(context.attributes.bank)] \(context.attributes.counterparty) → \(context.attributes.accountName)")
                    .font(.custom("Pretendard-Regular", size: 12)).foregroundStyle(.secondary).lineLimit(1)
            }
            Spacer()
            VStack(alignment: .trailing, spacing: 3) {
                Text(context.attributes.amountText).font(.custom("Pretendard-Bold", size: 17)).foregroundStyle(.blue).lineLimit(1).minimumScaleFactor(0.6)
                if !context.state.landed {
                    Text(timerInterval: Date()...context.state.fireDate, countsDown: true).monospacedDigit()
                        .font(.custom("Pretendard-Regular", size: 12)).foregroundStyle(.secondary).frame(width: 60, alignment: .trailing)
                }
            }
        }
        .padding(14)
    }
}

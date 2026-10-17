import SwiftUI

/// 인생 이벤트 timeline: 교통사고 · 경찰 단속 · 구속 · 실직, newest first, plus the current 부상/수감 status.
struct LifeEventTimelineCard: View {
    @EnvironmentObject var store: BankStore

    var body: some View {
        Card {
            VStack(alignment: .leading, spacing: 10) {
                HStack {
                    Text("인생 이벤트").font(.app(.headline, .semibold))
                    Spacer()
                    if store.life.heat > 0.5 {
                        Label("위험 수위 \(Int(min(99, store.life.heat * 25)))%", systemImage: "exclamationmark.triangle.fill")
                            .font(.app(.caption2, .semibold)).foregroundStyle(.orange).accessibilityIdentifier("life-heat")
                    }
                }
                if store.isJailed, let u = store.life.jailedUntil {
                    Label("수감 중 · \(u.formatted(date: .abbreviated, time: .omitted))까지 급여 없음", systemImage: "lock.fill")
                        .font(.app(.caption, .semibold)).foregroundStyle(.red)
                } else if store.isInjured, let u = store.life.injuredUntil {
                    Label("부상 회복 중 · \(u.formatted(date: .abbreviated, time: .omitted))까지 알바 불가", systemImage: "bandage.fill")
                        .font(.app(.caption, .semibold)).foregroundStyle(.orange)
                }
                if store.life.records.isEmpty {
                    Text("아직 아무 일도 없었어요. 살다 보면 사고도, 단속도, 실직도 옵니다.").font(.app(.subheadline)).foregroundStyle(.secondary)
                }
                ForEach(store.life.records.reversed().prefix(30)) { e in
                    VStack(alignment: .leading, spacing: 3) {
                        HStack {
                            Text("\(e.kind.emoji) \(e.title)").font(.app(.subheadline, .semibold))
                            Spacer()
                            Text(e.date.formatted(date: .abbreviated, time: .omitted)).font(.app(.caption2)).foregroundStyle(.tertiary)
                        }
                        Text(e.detail).font(.app(.caption)).foregroundStyle(.secondary)
                        if e.cash != 0 { MoneyText(amount: e.cash, signed: true, font: .app(.caption, .semibold), hideable: false).foregroundStyle(.secondary) }
                    }
                    .accessibilityElement(children: .contain)
                    .accessibilityIdentifier("life-event-\(e.kind.rawValue)")
                    Divider()
                }
            }
        }
    }
}

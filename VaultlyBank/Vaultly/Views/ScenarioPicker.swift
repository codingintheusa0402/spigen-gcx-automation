import SwiftUI

/// Pick a life to start from. Used on the second onboarding page and in 설정 > 시뮬레이션.
struct ScenarioPicker: View {
    @EnvironmentObject var store: BankStore
    @EnvironmentObject var shop: StoreService
    var onDone: (() -> Void)? = nil
    @State private var selected: Scenario?
    @State private var paywall = false

    private func gated(_ s: Scenario) -> Bool { (s == .heir || s == .lottery) && !shop.allows(.scenarios) }

    var body: some View {
        VStack(spacing: 12) {
            ForEach(Scenario.allCases) { s in
                Button { selected = s } label: {
                    HStack(alignment: .top, spacing: 14) {
                        Text(s.emoji).font(.app(size: 30))
                        VStack(alignment: .leading, spacing: 4) {
                            HStack {
                                Text(s.title).font(.app(.headline, .semibold))
                                Text("\(s.tier.emoji) \(s.tier.title)").font(.app(.caption)).foregroundStyle(.secondary)
                            }
                            Text(s.summary).font(.app(.caption)).foregroundStyle(.secondary).multilineTextAlignment(.leading)
                        }
                        Spacer()
                        if gated(s) { Image(systemName: "lock.fill").foregroundStyle(.tertiary) }
                        Image(systemName: selected == s ? "checkmark.circle.fill" : "circle").foregroundStyle(selected == s ? store.theme.accent : Color.secondary.opacity(0.4))
                    }
                    .padding(14)
                    .background(store.theme.card, in: RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
                    .overlay(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous).stroke(selected == s ? store.theme.accent : Color.primary.opacity(0.06)))
                }.buttonStyle(.plain)
            }
            Button {
                if let s = selected, gated(s) { paywall = true; return }
                if let s = selected { store.applyScenario(s); store.catchUp() }
                onDone?()
            } label: {
                Text(selected.map { "\($0.title)로 시작" } ?? "시나리오 선택").frame(maxWidth: .infinity).padding(.vertical, 6)
            }
            .buttonStyle(.borderedProminent).disabled(selected == nil)
            .accessibilityIdentifier("scenario-start")
        }
        .sheet(isPresented: $paywall) { ProPaywall() }
    }
}

struct ScenarioSettingsPage: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    var body: some View {
        ScrollView {
            VStack(spacing: 16) {
                Text("현재 계좌 잔액과 재력 설정을 시나리오에 맞게 바꾸고, 예약 입금을 추가합니다. 기존 거래 내역은 유지됩니다.")
                    .font(.app(.footnote)).foregroundStyle(.secondary).frame(maxWidth: .infinity, alignment: .leading)
                ScenarioPicker { dismiss() }
            }
            .padding()
        }
        .background(store.theme.background)
        .navigationTitle("시나리오")
        .navigationBarTitleDisplayMode(.inline)
    }
}

import SwiftUI
import StoreKit

/// Shown only in the store variant when a Pro feature is tapped.
struct ProPaywall: View {
    @EnvironmentObject var store: BankStore
    @EnvironmentObject var shop: StoreService
    @Environment(\.dismiss) private var dismiss
    @State private var busy = false

    var body: some View {
        NavigationStack {
            VStack(spacing: 20) {
                Image(systemName: "crown.fill").font(.app(size: 44, .bold)).foregroundStyle(store.theme.accent).padding(.top, 20)
                Text("Vaultly Pro").font(.app(.title, .bold))
                VStack(alignment: .leading, spacing: 10) {
                    ForEach([ProFeature.themes, .scenarios, .acceleration, .widgets, .challengeShare], id: \.rawValue) { f in
                        Label(f.title, systemImage: "checkmark.circle.fill").foregroundStyle(.primary)
                    }
                }
                .font(.app(.body))
                .padding().frame(maxWidth: .infinity, alignment: .leading)
                .background(store.theme.card, in: RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
                Spacer()
                if shop.products.isEmpty {
                    ProgressView().task { await shop.load() }
                }
                ForEach(shop.products, id: \.id) { p in
                    Button {
                        busy = true
                        Task { await shop.purchase(p); busy = false; if shop.isPro { dismiss() } }
                    } label: {
                        HStack {
                            Text(p.displayName).font(.app(.body, .semibold))
                            Spacer()
                            Text(p.displayPrice).font(.app(.body, .bold))
                        }
                        .padding(.vertical, 6)
                    }
                    .buttonStyle(.borderedProminent).disabled(busy)
                }
                Button("구매 복원") { Task { await shop.restore(); if shop.isPro { dismiss() } } }.font(.app(.subheadline))
                if let e = shop.lastError { Text(e).font(.app(.caption)).foregroundStyle(.red) }
                Text("모든 기능은 시뮬레이션이며 실제 금융 서비스가 아닙니다.").font(.app(.caption2)).foregroundStyle(.tertiary)
            }
            .padding(24)
            .background(store.theme.background)
            .toolbar { ToolbarItem(placement: .cancellationAction) { Button("닫기") { dismiss() } } }
        }
    }
}

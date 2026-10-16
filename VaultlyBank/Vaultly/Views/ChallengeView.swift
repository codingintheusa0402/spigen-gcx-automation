import SwiftUI

/// 무지출 챌린지: log what you resisted buying; the amount lands in a virtual 참은 돈 account.
struct ChallengeView: View {
    @EnvironmentObject var store: BankStore
    @EnvironmentObject var shop: StoreService
    @State private var logging = false
    @State private var paywall = false
    @State private var shareImage: Image?

    private var now: Date { store.now }
    private var c: ChallengeState { store.challenge }

    var body: some View {
        ScrollView {
            VStack(spacing: 16) {
                heroCard
                HStack(spacing: 12) {
                    statTile("🔥 연속", "\(c.streak(now: now))일", sub: "최고 \(c.bestStreak())일")
                    statTile("이번 달", "\(c.items(inMonthOf: now).count)번", sub: "참았어요")
                    statTile("누적", Money.format(c.totalSaved, code: store.currency), sub: "참은 돈")
                }
                Button { logging = true } label: {
                    Label("참았다! 기록하기", systemImage: "hand.raised.fill")
                        .font(.app(.body, .bold)).frame(maxWidth: .infinity).padding(.vertical, 14)
                }
                .buttonStyle(.borderedProminent)
                .accessibilityIdentifier("resist-button")
                historyCard
                achievementsCard
            }
            .padding()
        }
        .background(store.theme.background)
        .navigationTitle("무지출 챌린지")
        .navigationBarTitleDisplayMode(.inline)
        .toolbar {
            ToolbarItem(placement: .topBarTrailing) {
                if let shareImage, shop.allows(.challengeShare) {
                    ShareLink(item: shareImage, preview: SharePreview("이번 달 참은 돈", image: shareImage)) {
                        Image(systemName: "square.and.arrow.up")
                    }
                } else if shareImage != nil {
                    Button { paywall = true } label: { Image(systemName: "square.and.arrow.up") }
                }
            }
        }
        .sheet(isPresented: $logging) { ResistSheet() }
        .sheet(isPresented: $paywall) { ProPaywall() }
        .onAppear { renderShareCard() }
        .onChange(of: store.challenge) { _, _ in renderShareCard() }
    }

    private var heroCard: some View {
        Card {
            VStack(alignment: .leading, spacing: 8) {
                Text("\(now.formatted(.dateTime.month(.wide))) 참은 돈").font(.app(.subheadline)).foregroundStyle(.secondary)
                MoneyText(amount: c.total(inMonthOf: now), font: .app(size: 34, .semibold), hideable: false)
                Text(c.isActive ? "참을 때마다 '참은 돈' 계좌에 가상 입금됩니다. 진짜 돈처럼 쌓이는 걸 보세요."
                                : "사고 싶었지만 참은 것을 기록하면, 그 금액이 가상 계좌에 입금됩니다.")
                    .font(.app(.caption)).foregroundStyle(.secondary)
            }
        }
    }

    private func statTile(_ title: String, _ value: String, sub: String) -> some View {
        VStack(alignment: .leading, spacing: 4) {
            Text(title).font(.app(.caption)).foregroundStyle(.secondary)
            Text(value).font(.app(.headline, .semibold)).monospacedDigit().lineLimit(1).minimumScaleFactor(0.6)
            Text(sub).font(.app(.caption2)).foregroundStyle(.tertiary)
        }
        .frame(maxWidth: .infinity, alignment: .leading)
        .padding(12)
        .background(store.theme.card, in: RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
        .overlay(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous).stroke(Color.primary.opacity(0.06)))
    }

    private var historyCard: some View {
        Card {
            VStack(alignment: .leading, spacing: 10) {
                Text("기록").font(.app(.headline, .semibold))
                let items = c.items.sorted { $0.date > $1.date }
                if items.isEmpty {
                    Text("아직 기록이 없어요. 배달 앱을 닫았다면 지금 기록해 보세요.").font(.app(.subheadline)).foregroundStyle(.secondary)
                }
                ForEach(items.prefix(60)) { item in
                    HStack(spacing: 12) {
                        Text(item.kind.emoji).font(.app(.title3))
                        VStack(alignment: .leading, spacing: 2) {
                            Text(item.title).font(.app(.subheadline, .semibold))
                            Text(item.date.formatted(date: .abbreviated, time: .shortened)).font(.app(.caption)).foregroundStyle(.secondary)
                        }
                        Spacer()
                        MoneyText(amount: item.amount, signed: true, font: .app(.subheadline, .semibold), hideable: false).foregroundStyle(store.theme.accent)
                    }
                    .contextMenu { Button(role: .destructive) { store.deleteResist(item.id) } label: { Label("삭제", systemImage: "trash") } }
                    if item.id != items.prefix(60).last?.id { Divider() }
                }
            }
        }
    }

    private var achievementsCard: some View {
        Card {
            VStack(alignment: .leading, spacing: 10) {
                HStack { Text("업적").font(.app(.headline, .semibold)); Spacer(); NavigationLink("전체 보기") { AchievementsView() }.font(.app(.subheadline)) }
                let ids: [AchievementID] = [.firstResist, .streak7, .streak30, .saved100k, .saved1m]
                HStack(spacing: 10) {
                    ForEach(ids) { a in
                        let done = store.data.achievements[a.rawValue] != nil
                        VStack(spacing: 4) {
                            Image(systemName: a.symbol).font(.app(.title3)).foregroundStyle(done ? store.theme.accent : .secondary)
                            Text(a.title).font(.app(.caption2)).foregroundStyle(done ? .primary : .secondary).lineLimit(1).minimumScaleFactor(0.7)
                        }
                        .frame(maxWidth: .infinity).padding(.vertical, 8)
                        .background(Color.primary.opacity(done ? 0.06 : 0.02), in: RoundedRectangle(cornerRadius: 10))
                        .opacity(done ? 1 : 0.6)
                    }
                }
            }
        }
    }

    @MainActor private func renderShareCard() {
        let card = ChallengeShareCard(bank: store.brand.bank, month: now, saved: c.total(inMonthOf: now), count: c.items(inMonthOf: now).count,
                                      streak: c.streak(now: now), total: c.totalSaved, currency: store.currency, accent: store.theme.accent)
        let renderer = ImageRenderer(content: card.frame(width: 360, height: 450))
        renderer.scale = 3
        if let ui = renderer.uiImage { shareImage = Image(uiImage: ui) }
    }
}

/// Square-ish card for sharing (Instagram story friendly ratio).
struct ChallengeShareCard: View {
    let bank: String
    let month: Date
    let saved: Decimal
    let count: Int
    let streak: Int
    let total: Decimal
    let currency: String
    let accent: Color

    var body: some View {
        ZStack {
            LinearGradient(colors: [accent, accent.opacity(0.75)], startPoint: .topLeading, endPoint: .bottomTrailing)
            VStack(alignment: .leading, spacing: 14) {
                Text("무지출 챌린지").font(.app(.subheadline, .semibold)).foregroundStyle(.white.opacity(0.85))
                Text("\(month.formatted(.dateTime.month(.wide))) 참은 돈").font(.app(.title3)).foregroundStyle(.white.opacity(0.9))
                Text(Money.format(saved, code: currency)).font(.app(size: 40, .bold)).foregroundStyle(.white).monospacedDigit().minimumScaleFactor(0.5).lineLimit(1)
                Spacer()
                HStack(spacing: 20) {
                    stat("\(count)번", "참음")
                    stat("\(streak)일", "연속")
                    stat(Money.format(total, code: currency), "누적")
                }
                Spacer()
                HStack {
                    Text(bank).font(.app(.caption, .semibold)).foregroundStyle(.white.opacity(0.8))
                    Spacer()
                    Text("Vaultly · 시뮬레이션").font(.app(.caption2)).foregroundStyle(.white.opacity(0.6))
                }
            }
            .padding(28)
        }
    }

    private func stat(_ v: String, _ t: String) -> some View {
        VStack(alignment: .leading, spacing: 2) {
            Text(v).font(.app(.headline, .bold)).foregroundStyle(.white).lineLimit(1).minimumScaleFactor(0.6)
            Text(t).font(.app(.caption)).foregroundStyle(.white.opacity(0.75))
        }
    }
}

struct ResistSheet: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    @State private var kind: ResistKind = .delivery
    @State private var title = ""
    @State private var amount: Decimal = ResistKind.delivery.defaultAmount
    @State private var saved: ResistedPurchase?

    var body: some View {
        NavigationStack {
            Form {
                Section("뭘 참았나요?") {
                    LazyVGrid(columns: Array(repeating: GridItem(.flexible()), count: 4), spacing: 10) {
                        ForEach(ResistKind.allCases) { k in
                            Button {
                                kind = k; amount = k.defaultAmount
                            } label: {
                                VStack(spacing: 4) {
                                    Text(k.emoji).font(.app(.title2))
                                    Text(k.title).font(.app(.caption2)).foregroundStyle(.primary)
                                }
                                .frame(maxWidth: .infinity).padding(.vertical, 8)
                                .background(kind == k ? store.theme.accent.opacity(0.18) : Color.primary.opacity(0.05), in: RoundedRectangle(cornerRadius: 10))
                                .overlay(RoundedRectangle(cornerRadius: 10).stroke(kind == k ? store.theme.accent : .clear))
                            }.buttonStyle(.plain)
                        }
                    }
                    .listRowInsets(EdgeInsets(top: 8, leading: 12, bottom: 8, trailing: 12))
                    TextField(kind.placeholder, text: $title)
                    AmountField(title: "안 쓴 금액", amount: $amount)
                }
                Section {
                    Text("기록하면 '참은 돈' 계좌에 \(Money.format(amount, code: store.currency))이 가상 입금됩니다.")
                        .font(.app(.footnote)).foregroundStyle(.secondary)
                }
            }
            .navigationTitle("참았다!")
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .cancellationAction) { Button("취소") { dismiss() } }
                ToolbarItem(placement: .confirmationAction) {
                    Button("기록") { saved = store.logResist(kind: kind, title: title, amount: amount) }.disabled(amount <= 0)
                }
            }
            .overlay {
                if let saved {
                    VStack(spacing: 12) {
                        SuccessCheck()
                        Text("+\(Money.format(saved.amount, code: store.currency))").font(.app(.title2, .bold))
                        Text("참은 돈 계좌에 입금됐어요").foregroundStyle(.secondary)
                    }
                    .frame(maxWidth: .infinity, maxHeight: .infinity).background(.regularMaterial)
                    .task { try? await Task.sleep(for: .seconds(1.2)); dismiss() }
                }
            }
        }
    }
}

struct AchievementsView: View {
    @EnvironmentObject var store: BankStore
    var body: some View {
        List {
            ForEach(AchievementID.allCases) { a in
                let date = store.data.achievements[a.rawValue]
                HStack(spacing: 14) {
                    Image(systemName: a.symbol).font(.app(.title2)).foregroundStyle(date != nil ? store.theme.accent : .secondary).frame(width: 32)
                    VStack(alignment: .leading, spacing: 2) {
                        Text(a.title).font(.app(.body, .semibold))
                        Text(a.detail).font(.app(.caption)).foregroundStyle(.secondary)
                    }
                    Spacer()
                    if let date { Text(date.formatted(date: .abbreviated, time: .omitted)).font(.app(.caption2)).foregroundStyle(.tertiary) }
                    else { Image(systemName: "lock.fill").foregroundStyle(.tertiary) }
                }
                .opacity(date != nil ? 1 : 0.6)
            }
        }
        .navigationTitle("업적")
    }
}

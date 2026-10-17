import SwiftUI

/// 내 방 — the opening screen: a Cyworld-style 미니룸 that reflects the player's wealth. The room itself is the tier
/// (고시원 → 원룸 → 아파트 → 펜트하우스 → 저택), furniture appears as net worth grows toward the next tier, clutter
/// shows up when the account is empty or in debt, and the pixel character lives its routine inside.
struct RoomView: View {
    @EnvironmentObject var store: BankStore
    @State private var endingShown = false

    var body: some View {
        NavigationStack {
            ScrollView {
                VStack(spacing: 14) {
                    header
                    RoomScene()
                    bubble
                    progressCard
                    menu
                }
                .padding(16)
            }
            .background(store.theme.background.ignoresSafeArea())
            .navigationTitle("미니룸")
            .navigationBarTitleDisplayMode(.inline)
            .homeRoutes()
            .sheet(item: $store.questResult) { OutcomeSheet(result: $0).presentationDetents([.medium]) }
            .sheet(isPresented: Binding(get: { store.game.isEnded && !endingShown }, set: { _ in endingShown = true })) { EndingView() }
        }
    }

    private var tierRoom: RoomManifest.Tier? { RoomAssets.shared.manifest?.tiers[store.currentTier.rawValue] }

    private var header: some View {
        let today = store.recentTransactions.filter { Calendar.current.isDate($0.date, inSameDayAs: store.now) }.count
        return Card {
            VStack(alignment: .leading, spacing: 6) {
                HStack(alignment: .firstTextBaseline) {
                    Text("\(store.settings.ownerName)님의 미니룸").font(.app(.title3, .bold))
                    MockChip()
                    Spacer()
                    Text("TODAY \(today) | TOTAL \(store.data.transactions.count)")
                        .font(.system(.caption2, design: .monospaced)).foregroundStyle(.secondary)
                }
                HStack(spacing: 8) {
                    Text("\(store.currentTier.emoji) \(store.currentTier.title)").font(.app(.subheadline, .semibold))
                    Text("·").foregroundStyle(.tertiary)
                    Text(tierRoom?.name ?? "방").font(.app(.subheadline, .semibold)).foregroundStyle(store.theme.accent).accessibilityIdentifier("room-name")
                    Spacer()
                    MoneyText(amount: store.totalAssets, font: .app(.subheadline, .semibold))
                }
            }
        }
    }

    /// Speech bubble: what the character is doing, holidays, a pending choice.
    @ViewBuilder private var bubble: some View {
        let slot = store.characterNow?.slot
        VStack(alignment: .leading, spacing: 8) {
            HStack(spacing: 8) {
                Text("💬").font(.app(.title3))
                VStack(alignment: .leading, spacing: 2) {
                    Text(store.isJailed ? "구치소에 있어요…" : store.isInjured ? "다쳐서 쉬는 중" : (slot.map { "\($0.label) 중" } ?? "쉬는 중")).font(.app(.subheadline, .semibold))
                        .accessibilityIdentifier("room-status")
                    if let h = KoreanHolidays.holiday(on: store.now) { Text("🎌 오늘은 \(h.name)").font(.app(.caption)).foregroundStyle(.red) }
                    else { Text(store.now.formatted(.dateTime.month().day().weekday(.wide).hour().minute().locale(Locale(identifier: "ko_KR")))).font(.app(.caption)).foregroundStyle(.secondary) }
                }
                Spacer()
            }
            if store.pendingQuest != nil {
                NavigationLink(value: HomeRoute.life) {
                    HStack {
                        Text("📩 새로운 선택이 왔어요").font(.app(.subheadline, .semibold))
                        Spacer()
                        Image(systemName: "chevron.right").font(.app(.caption))
                    }
                    .padding(10)
                    .background(store.theme.accent.opacity(0.12), in: RoundedRectangle(cornerRadius: 10))
                }
                .buttonStyle(.plain)
                .accessibilityIdentifier("room-quest")
            }
        }
        .padding(12)
        .frame(maxWidth: .infinity, alignment: .leading)
        .background(store.theme.card, in: RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
    }

    private var progressCard: some View {
        let state = store.roomState
        let total = tierRoom?.props.filter { !$0.isClutter }.count ?? 0
        let have = tierRoom.map { state.visible($0.props).filter { !$0.isClutter }.count } ?? 0
        return Card {
            VStack(alignment: .leading, spacing: 8) {
                HStack {
                    Text("가구 \(have)/\(total)").font(.app(.subheadline, .semibold))
                    Spacer()
                    if let next = state.nextTier, let room = RoomAssets.shared.manifest?.tiers[next.rawValue] {
                        Text("다음 방 \(room.name) · \(Int(state.progress * 100))%").font(.app(.caption)).foregroundStyle(.secondary)
                    } else {
                        Text("최고의 방 · \(Int(state.progress * 100))%").font(.app(.caption)).foregroundStyle(.secondary)
                    }
                }
                GeometryReader { geo in
                    Capsule().fill(Color.primary.opacity(0.07)).overlay(alignment: .leading) { Capsule().fill(store.theme.accent).frame(width: geo.size.width * state.progress) }
                }.frame(height: 6)
                Text(state.poor ? "통장이 비어 있어요 — 고지서와 빈 병이 쌓입니다. 순자산이 늘면 가구가 하나씩 들어와요."
                                : "순자산이 늘수록 가구가 하나씩 들어오고, 승급하면 더 좋은 방으로 이사해요.")
                    .font(.app(.caption2)).foregroundStyle(.tertiary)
            }
        }
    }

    private var menu: some View {
        HStack(spacing: 10) {
            menuTile("내 인생", "figure.walk", .life, id: "room-menu-life")
            menuTile("알바", "briefcase.fill", .jobs, id: "room-menu-jobs")
            menuTile("일정", "calendar", .schedule, id: "room-menu-schedule")
            menuTile("챌린지", "flame.fill", .challenge, id: "room-menu-challenge")
        }
    }

    private func menuTile(_ title: String, _ symbol: String, _ route: HomeRoute, id: String) -> some View {
        NavigationLink(value: route) {
            VStack(spacing: 6) {
                Image(systemName: symbol).font(.app(.title3)).foregroundStyle(store.theme.accent)
                Text(title).font(.app(.caption, .semibold))
            }
            .frame(maxWidth: .infinity).padding(.vertical, 12)
            .background(store.theme.card, in: RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
        }
        .buttonStyle(.plain)
        .accessibilityIdentifier(id)
    }
}

/// The room canvas: background, furniture and the character, y-sorted so things in front overlap things behind.
struct RoomScene: View {
    @EnvironmentObject var store: BankStore

    private enum Layer { case prop(RoomManifest.Prop), character }

    var body: some View {
        if let m = RoomAssets.shared.manifest, let room = m.tiers[store.currentTier.rawValue] {
            let state = store.roomState
            let props = state.visible(room.props)
            GeometryReader { geo in
                let s = geo.size.width / CGFloat(m.canvas[0])
                let cs = CGFloat(room.scale) * s
                let anchorY = room.anchor[1]
                let layers: [(Int, Layer)] = props.map { ($0.bottom, Layer.prop($0)) } + [(anchorY, .character)]
                ZStack(alignment: .topLeading) {
                    pixel(room.bg).frame(width: geo.size.width, height: CGFloat(m.canvas[1]) * s)
                    ForEach(Array(layers.sorted { $0.0 < $1.0 }.enumerated()), id: \.offset) { _, layer in
                        switch layer.1 {
                        case .prop(let p):
                            pixel("rp_\(p.name)")
                                .frame(width: CGFloat(p.w) * s, height: CGFloat(p.h) * s)
                                .offset(x: CGFloat(p.x) * s, y: CGFloat(p.y) * s)
                                .accessibilityLabel(p.name)
                                .accessibilityIdentifier("room-prop-\(p.name)")
                        case .character:
                            if let c = store.characterNow {
                                CharacterSpriteView(look: c.look, activity: c.slot.activity, scale: cs,
                                                    showProps: ["sleep", "injured"].contains(c.slot.activity))
                                    .offset(x: CGFloat(room.anchor[0]) * s - 48 * cs, y: CGFloat(anchorY) * s - 68 * cs)
                            }
                        }
                    }
                }
            }
            .aspectRatio(CGFloat(m.canvas[0]) / CGFloat(m.canvas[1]), contentMode: .fit)
            .clipShape(RoundedRectangle(cornerRadius: 14, style: .continuous))
            .overlay(RoundedRectangle(cornerRadius: 14, style: .continuous).stroke(Color.black.opacity(0.35), lineWidth: 2))
            .accessibilityElement(children: .contain)
            .accessibilityIdentifier("room-scene")
        } else {
            Card { Text("미니룸을 불러올 수 없어요").foregroundStyle(.secondary) }
        }
    }

    private func pixel(_ name: String) -> some View {
        Image(uiImage: UIImage(named: name) ?? UIImage()).interpolation(.none).resizable()
    }
}

extension BankStore {
    /// Room = tier + progress toward the next tier + whether the account is empty / in debt.
    var roomState: RoomState {
        let balance = payAccount?.balance ?? 0
        let poor = balance <= 0 || (data.lifestyle.debtAccountID != nil && totalDebt > balance)
        return RoomState(tier: currentTier, progress: RoomState.progress(netWorth: netWorth, tier: currentTier), poor: poor)
    }
}

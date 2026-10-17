import SwiftUI
import LocalAuthentication

struct RootView: View {
    @EnvironmentObject var store: BankStore
    @EnvironmentObject var shop: StoreService
    @EnvironmentObject var ads: AdService
    @Environment(\.scenePhase) private var scenePhase
    @State private var locked = false
    @State private var unlocking = false

    var body: some View {
        ZStack(alignment: .top) {
            TabView {
                RoomView().tabItem { Label("내 방", systemImage: "house.fill") }
                HomeView().tabItem { Label("은행", systemImage: "building.columns.fill") }
                ProductsView().tabItem { Label("상품", systemImage: "chart.line.uptrend.xyaxis") }
                SettingsView().tabItem { Label("설정", systemImage: "gearshape.fill") }
            }
            if AppVariant.isStore { MockWatermark() }
            if let banner = store.banner {
                BannerView(message: banner)
                    .transition(.move(edge: .top).combined(with: .opacity))
                    .onTapGesture { store.banner = nil }
                    .task(id: banner.id) {
                        try? await Task.sleep(for: .seconds(4))
                        if store.banner?.id == banner.id { store.banner = nil }
                    }
            }
            if locked { LockScreen(unlock: unlock).transition(.opacity) }
            if let ad = ads.fake { FakeAdOverlay(ad: ad).zIndex(10) }
        }
        .onAppear {
            if store.settings.appLock && store.settings.hasOnboarded { locked = true; unlock() }
            if store.settings.hasOnboarded { ads.start() }
            store.grantProBonusIfNeeded(isPro: shop.isPro)
        }
        .onChange(of: store.settings.hasOnboarded) { _, on in if on { ads.start() } }
        .onChange(of: shop.isPro) { _, pro in store.grantProBonusIfNeeded(isPro: pro) }
        .onChange(of: scenePhase) { _, phase in
            guard store.settings.appLock, store.settings.hasOnboarded else { return }
            if phase == .background { locked = true }
            if phase == .active, locked { unlock() }
        }
        .animation(.spring(duration: 0.4), value: store.banner)
        .themed(store.theme, appearance: store.settings.appearance)
        .sheet(isPresented: Binding(get: { !store.settings.hasOnboarded }, set: { _ in })) {
            OnboardingView().interactiveDismissDisabled()
        }
    }

    private func unlock() {
        guard !unlocking else { return }
        unlocking = true
        let ctx = LAContext()
        ctx.localizedCancelTitle = "취소"
        var err: NSError?
        guard ctx.canEvaluatePolicy(.deviceOwnerAuthentication, error: &err) else { locked = false; unlocking = false; return }
        ctx.evaluatePolicy(.deviceOwnerAuthentication, localizedReason: "앱 잠금을 해제합니다") { ok, _ in
            Task { @MainActor in
                withAnimation { locked = !ok }
                unlocking = false
            }
        }
    }
}

struct LockScreen: View {
    @EnvironmentObject var store: BankStore
    let unlock: () -> Void
    var body: some View {
        ZStack {
            store.theme.background.ignoresSafeArea()
            VStack(spacing: 18) {
                Image(systemName: "lock.fill").font(.app(size: 44, .bold)).foregroundStyle(store.theme.accent)
                Text(store.brand.bank).font(.app(.title2, .bold))
                Text("잠금 상태입니다").foregroundStyle(.secondary)
                Button { unlock() } label: {
                    Label("Face ID로 잠금 해제", systemImage: "faceid").frame(maxWidth: .infinity).padding(.vertical, 6)
                }
                .buttonStyle(.borderedProminent)
                .padding(.horizontal, 40)
            }
        }
    }
}

struct OnboardingView: View {
    @EnvironmentObject var store: BankStore
    @State private var name = ""
    @State private var pickingScenario = false
    @State private var startTier: WealthTier = .working     // 서민 or 거지 — the player's call

    var body: some View {
        if pickingScenario { scenarioPage } else { welcomePage }
    }

    private var scenarioPage: some View {
        ScrollView {
            VStack(spacing: 16) {
                Text("어떤 삶으로 시작할까요?").font(.app(.title2, .bold)).padding(.top, 20)
                Text("잔액·재력·예약 입금이 시나리오에 맞게 설정됩니다. 나중에 설정에서 바꿀 수 있어요.").font(.app(.footnote)).foregroundStyle(.secondary).multilineTextAlignment(.center)
                ScenarioPicker { finish() }
                Button("직접 설정할게요") { finish() }.font(.app(.subheadline)).padding(.top, 4)
            }
            .padding(24)
        }
    }

    private func finish() {
        store.update { $0.ownerName = name.isEmpty ? $0.ownerName : name; $0.hasOnboarded = true }
        // No scenario picked → start as the tier chosen on the welcome page (a scenario sets its own tier first).
        if store.lifestyle.tier == nil {
            store.applyTier(startTier, accountID: store.checkingAccounts.first?.id, resetBalance: true, backfillDays: 30)
        }
        store.requestNotifications()
    }

    private var welcomePage: some View {
        ScrollView {
        VStack(spacing: 24) {
            Image(systemName: "building.columns.circle.fill").font(.app(size: 72)).foregroundStyle(store.theme.accent).padding(.top, 24)
            Text("Vaultly").font(.app(.largeTitle, .bold))
            Text(AppVariant.isStore ? "가상 자산으로 즐기는 인생 시뮬레이션 게임" : "가상 자산으로 즐기는 뱅킹 시뮬레이터")
                .foregroundStyle(.secondary)
            VStack(alignment: .leading, spacing: 10) {
                Label("모든 계좌와 거래는 가상이며 기기에만 저장됩니다.", systemImage: "iphone")
                Label("실제 은행·결제 서비스와 연결되지 않습니다.", systemImage: "wifi.slash")
                Label("사고 싶은 걸 참으면 '참은 돈'이 가상으로 쌓입니다.", systemImage: "flame")
            }
            .font(.app(.footnote)).foregroundStyle(.secondary)
            .padding().background(Color.primary.opacity(0.05), in: RoundedRectangle(cornerRadius: 16))
            TextField("이름 (예: 홍길동)", text: $name)
                .textFieldStyle(.roundedBorder)
                .submitLabel(.done)
            VStack(alignment: .leading, spacing: 8) {
                Text("어디서 시작할까요?").font(.app(.subheadline, .semibold))
                HStack(spacing: 10) {
                    startChoice(.working, emoji: "🧑‍💼", title: "서민", detail: "월급 받는 직장인\n원룸 · 잔액 \(Money.format(WealthTier.working.startingBalance, code: store.currency))", id: "start-working")
                    startChoice(.broke, emoji: "🥲", title: "거지", detail: "일용직·배달로 하루살이\n고시원 · 빚 480만원", id: "start-broke")
                }
            }
        }
        .padding(28)
        }
        .scrollDismissesKeyboard(.interactively)
        .safeAreaInset(edge: .bottom) {           // stays above the keyboard, so 시작하기 is always reachable
            VStack(spacing: 10) {
                Button { finish() } label: {
                    Text("시작하기").frame(maxWidth: .infinity).padding(.vertical, 6)
                }
                .buttonStyle(.borderedProminent)
                Button("시나리오로 시작 (월급날 · 로또 · 재벌 2세 · 퇴사)") { pickingScenario = true }
                    .font(.app(.subheadline))
            }
            .padding(.horizontal, 28).padding(.vertical, 12)
            .background(.bar)
        }
    }

    private func startChoice(_ tier: WealthTier, emoji: String, title: String, detail: String, id: String) -> some View {
        let on = startTier == tier
        return Button { startTier = tier } label: {
            VStack(alignment: .leading, spacing: 4) {
                HStack { Text(emoji); Text(title).font(.app(.headline, .bold)); Spacer(); if on { Image(systemName: "checkmark.circle.fill").foregroundStyle(store.theme.accent) } }
                Text(detail).font(.app(.caption2)).foregroundStyle(.secondary).multilineTextAlignment(.leading)
            }
            .padding(12).frame(maxWidth: .infinity, alignment: .leading)
            .background(on ? store.theme.accent.opacity(0.12) : Color.primary.opacity(0.05), in: RoundedRectangle(cornerRadius: 12))
            .overlay(RoundedRectangle(cornerRadius: 12).stroke(on ? store.theme.accent : .clear, lineWidth: 1.5))
        }
        .buttonStyle(.plain)
        .accessibilityIdentifier(id)
        .accessibilityAddTraits(on ? .isSelected : [])
    }
}

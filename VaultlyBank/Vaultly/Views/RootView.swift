import SwiftUI
import LocalAuthentication

struct RootView: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.scenePhase) private var scenePhase
    @State private var locked = false
    @State private var unlocking = false

    var body: some View {
        ZStack(alignment: .top) {
            TabView {
                HomeView().tabItem { Label("홈", systemImage: "house.fill") }
                InvestView().tabItem { Label("증권", systemImage: "chart.xyaxis.line") }
                ProductsView().tabItem { Label("상품", systemImage: "chart.line.uptrend.xyaxis") }
                ScheduleView().tabItem { Label("예약 알림", systemImage: "bell.badge") }
                SettingsView().tabItem { Label("설정", systemImage: "gearshape.fill") }
            }
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
        }
        .onAppear { if store.settings.appLock && store.settings.hasOnboarded { locked = true; unlock() } }
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

    var body: some View {
        VStack(spacing: 24) {
            Spacer()
            Image(systemName: "building.columns.circle.fill").font(.app(size: 72)).foregroundStyle(store.theme.accent)
            Text("Vaultly").font(.app(.largeTitle, .bold))
            Text("가상 자산으로 즐기는 뱅킹 시뮬레이터")
                .foregroundStyle(.secondary)
            VStack(alignment: .leading, spacing: 10) {
                Label("모든 계좌와 거래는 가상이며 기기에만 저장됩니다.", systemImage: "iphone")
                Label("실제 은행·결제 서비스와 연결되지 않습니다.", systemImage: "wifi.slash")
                Label("화면의 SIM 배지는 시뮬레이션 표시입니다.", systemImage: "checkmark.shield")
            }
            .font(.app(.footnote)).foregroundStyle(.secondary)
            .padding().background(Color.primary.opacity(0.05), in: RoundedRectangle(cornerRadius: 16))
            TextField("이름 (예: 홍길동)", text: $name)
                .textFieldStyle(.roundedBorder)
            Spacer()
            Button {
                store.update { $0.ownerName = name.isEmpty ? $0.ownerName : name; $0.hasOnboarded = true }
                store.requestNotifications()
            } label: {
                Text("시작하기").frame(maxWidth: .infinity).padding(.vertical, 6)
            }
            .buttonStyle(.borderedProminent)
        }
        .padding(28)
    }
}

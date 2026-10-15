import SwiftUI

struct RootView: View {
    @EnvironmentObject var store: BankStore

    var body: some View {
        ZStack(alignment: .top) {
            TabView {
                HomeView().tabItem { Label("홈", systemImage: "house.fill") }
                TransferView().tabItem { Label("송금", systemImage: "paperplane.fill") }
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
        }
        .animation(.spring(duration: 0.4), value: store.banner)
        .themed(store.theme)
        .sheet(isPresented: Binding(get: { !store.settings.hasOnboarded }, set: { _ in })) {
            OnboardingView().interactiveDismissDisabled()
        }
    }
}

struct OnboardingView: View {
    @EnvironmentObject var store: BankStore
    @State private var name = ""

    var body: some View {
        VStack(spacing: 24) {
            Spacer()
            Image(systemName: "building.columns.circle.fill").font(.system(size: 72)).foregroundStyle(store.theme.accent)
            Text("Vaultly").font(.largeTitle.bold())
            Text("가상 자산으로 즐기는 뱅킹 시뮬레이터")
                .foregroundStyle(.secondary)
            VStack(alignment: .leading, spacing: 10) {
                Label("모든 계좌와 거래는 가상이며 기기에만 저장됩니다.", systemImage: "iphone")
                Label("실제 은행·결제 서비스와 연결되지 않습니다.", systemImage: "wifi.slash")
                Label("화면의 SIM 배지는 시뮬레이션 표시입니다.", systemImage: "checkmark.shield")
            }
            .font(.footnote).foregroundStyle(.secondary)
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

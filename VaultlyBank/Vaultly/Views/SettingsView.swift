import SwiftUI

struct SettingsView: View {
    @EnvironmentObject var store: BankStore
    @State private var confirmReset = false
    @State private var name = ""

    var body: some View {
        NavigationStack {
            Form {
                Section("프로필") {
                    TextField("이름", text: $name)
                        .onAppear { name = store.settings.ownerName }
                        .onSubmit { store.update { $0.ownerName = name } }
                }
                Section("테마 (가상 브랜드)") {
                    ForEach(Theme.all) { t in
                        Button { store.update { $0.themeID = t.id } } label: {
                            HStack(spacing: 12) {
                                RoundedRectangle(cornerRadius: 8)
                                    .fill(LinearGradient(colors: t.cardGradient, startPoint: .topLeading, endPoint: .bottomTrailing))
                                    .frame(width: 48, height: 32)
                                VStack(alignment: .leading) {
                                    Text(t.bankName).fontWeight(.semibold)
                                    Text(t.tagline).font(.caption).foregroundStyle(.secondary)
                                }
                                Spacer()
                                if store.settings.themeID == t.id { Image(systemName: "checkmark").foregroundStyle(store.theme.accent) }
                            }
                        }.buttonStyle(.plain)
                    }
                }
                Section("표시") {
                    Picker("통화", selection: Binding(get: { store.settings.currencyCode }, set: { c in store.update { $0.currencyCode = c } })) {
                        Text("원 (KRW)").tag("KRW"); Text("달러 (USD)").tag("USD"); Text("유로 (EUR)").tag("EUR"); Text("엔 (JPY)").tag("JPY")
                    }
                    Toggle("잔액 숨기기", isOn: Binding(get: { store.settings.hideBalances }, set: { v in store.update { $0.hideBalances = v } }))
                }
                Section {
                    Picker("이자소득세", selection: Binding(get: { store.settings.interestTaxRate }, set: { v in store.update { $0.interestTaxRate = v } })) {
                        Text("15.4% (일반)").tag(0.154); Text("9.5% (세금우대)").tag(0.095); Text("비과세").tag(0.0)
                    }
                } header: { Text("이자") } footer: { Text("만기 이자 지급 시 적용됩니다.") }
                Section("알림") {
                    HStack {
                        Text("알림 권한"); Spacer()
                        Text(store.notificationsAuthorized ? "허용됨" : "없음").foregroundStyle(.secondary)
                    }
                    if !store.notificationsAuthorized { Button("알림 권한 요청") { store.requestNotifications() } }
                    Button("알림 다시 동기화") { store.syncNotifications() }
                }
                Section {
                    Button("모든 데이터 초기화", role: .destructive) { confirmReset = true }
                }
                Section("안내") {
                    VStack(alignment: .leading, spacing: 8) {
                        Label("이 앱은 뱅킹 시뮬레이터입니다.", systemImage: "info.circle")
                        Text("표시되는 모든 계좌, 잔액, 거래, 이자, 알림은 사용자가 설정한 가상의 데이터이며 실제 금융기관·결제 서비스와 연결되어 있지 않습니다. 실제 자금의 이동이나 금융 거래는 발생하지 않습니다.")
                            .font(.caption).foregroundStyle(.secondary)
                        Text("Nova Bank · Heritage Trust · Orbit Pay 는 이 앱을 위해 만든 가상의 브랜드입니다.")
                            .font(.caption).foregroundStyle(.secondary)
                    }
                }
            }
            .navigationTitle("설정")
            .toolbar { ToolbarItem(placement: .topBarLeading) { SimBadge() } }
            .confirmationDialog("모든 계좌·거래·예약을 삭제하고 처음 상태로 되돌릴까요?", isPresented: $confirmReset, titleVisibility: .visible) {
                Button("초기화", role: .destructive) { store.resetAll() }
            }
        }
    }
}

import SwiftUI

/// "전체 메뉴" style settings, laid out the way Korean bank apps do it:
/// profile → 보안·알림 → 화면 → 금융사 이름 → 시뮬레이션 → 고객센터·앱 정보.
struct SettingsView: View {
    @EnvironmentObject var store: BankStore
    @State private var confirmReset = false
    @State private var showWealth = false
    @State private var editCash = false

    var body: some View {
        NavigationStack {
            Form {
                profileHeader

                Section("보안 · 알림") {
                    Toggle(isOn: Binding(get: { store.settings.appLock }, set: { v in store.update { $0.appLock = v } })) {
                        Label("앱 잠금 (Face ID · 암호)", systemImage: "faceid")
                    }
                    Toggle(isOn: Binding(get: { store.settings.hideBalances }, set: { v in store.update { $0.hideBalances = v } })) {
                        Label("잔액 숨기기", systemImage: "eye.slash")
                    }
                    HStack {
                        Label("입금 알림", systemImage: "bell.badge")
                        Spacer()
                        if store.notificationsAuthorized {
                            Text("허용됨").foregroundStyle(.secondary)
                        } else {
                            Button("허용하기") { store.requestNotifications() }
                        }
                    }
                }

                Section("화면") {
                    NavigationLink { ThemePicker() } label: {
                        HStack { Label("테마", systemImage: "paintpalette"); Spacer(); Text(store.theme.displayName).foregroundStyle(.secondary) }
                    }
                    Picker(selection: Binding(get: { store.settings.appearance }, set: { v in store.update { $0.appearance = v } })) {
                        ForEach(Appearance.allCases) { Text($0.title).tag($0) }
                    } label: { Label("화면 모드", systemImage: "circle.lefthalf.filled") }
                    Picker(selection: Binding(get: { store.settings.currencyCode }, set: { c in store.update { $0.currencyCode = c } })) {
                        Text("원 (KRW)").tag("KRW"); Text("달러 (USD)").tag("USD"); Text("유로 (EUR)").tag("EUR"); Text("엔 (JPY)").tag("JPY")
                    } label: { Label("통화", systemImage: "wonsign.circle") }
                }

                Section {
                    nameField("은행", symbol: "building.columns", keyPath: \.bankName, placeholder: store.theme.bankName)
                    nameField("증권사", symbol: "chart.xyaxis.line", keyPath: \.brokerName, placeholder: store.theme.brokerName)
                    nameField("카드사", symbol: "creditcard", keyPath: \.cardName, placeholder: Brand.defaultCard)
                    nameField("캐피탈", symbol: "percent", keyPath: \.capitalName, placeholder: Brand.defaultCapital)
                    Button("기본값으로 되돌리기") {
                        store.update { $0.bankName = ""; $0.brokerName = ""; $0.cardName = ""; $0.capitalName = "" }
                    }
                } header: { Text("금융사 이름") } footer: {
                    Text("화면과 거래 내역에 표시되는 이름입니다. 비워두면 테마 기본 이름을 씁니다. 모두 가상의 표기이며 실제 금융기관과 무관합니다.")
                }

                Section {
                    Button { showWealth = true } label: {
                        HStack {
                            Label("재력 설정", systemImage: "person.crop.circle.badge.checkmark")
                            Spacer()
                            Text(store.lifestyle.tier.map { "\($0.emoji) \($0.title)" } ?? "꺼짐").foregroundStyle(.secondary)
                        }
                    }
                    NavigationLink { TimeFlowSettings() } label: {
                        HStack {
                            Label("시간 흐름", systemImage: "forward.fill")
                            Spacer()
                            Text(TimeSpeed.allCases.first { $0.rawValue == store.clock.speed }?.short ?? "\(Int(store.clock.speed))x").foregroundStyle(.secondary)
                        }
                    }
                    Toggle(isOn: Binding(get: { store.portfolio.simulatePrices }, set: { v in store.updatePortfolio { $0.simulatePrices = v; $0.lastTick = Date() } })) {
                        Label("시세 자동 변동", systemImage: "waveform.path.ecg")
                    }
                    Button { editCash = true } label: { Label("예수금 · 환율 설정", systemImage: "dollarsign.circle") }
                    Picker(selection: Binding(get: { store.settings.interestTaxRate }, set: { v in store.update { $0.interestTaxRate = v } })) {
                        Text("15.4% (일반)").tag(0.154); Text("9.5% (세금우대)").tag(0.095); Text("비과세").tag(0.0)
                    } label: { Label("이자소득세", systemImage: "percent") }
                    Button(role: .destructive) { confirmReset = true } label: { Label("모든 데이터 초기화", systemImage: "trash") }
                } header: { Text("시뮬레이션") }

                Section("고객센터 · 앱 정보") {
                    NavigationLink { InfoPage(title: "공지사항", items: InfoPage.notices) } label: { Label("공지사항", systemImage: "megaphone") }
                    NavigationLink { InfoPage(title: "자주 묻는 질문", items: InfoPage.faq) } label: { Label("자주 묻는 질문", systemImage: "questionmark.circle") }
                    NavigationLink { InfoPage(title: "이용안내 · 개인정보", items: InfoPage.terms) } label: { Label("이용안내 · 개인정보", systemImage: "doc.text") }
                    HStack {
                        Label("앱 버전", systemImage: "info.circle")
                        Spacer()
                        Text(Self.version).foregroundStyle(.secondary)
                    }
                }
            }
            .navigationTitle("설정")
            .toolbar { ToolbarItem(placement: .topBarLeading) { SimBadge() } }
            .sheet(isPresented: $showWealth) { WealthView() }
            .sheet(isPresented: $editCash) { CashEditorSheet() }
            .confirmationDialog("모든 계좌·거래·예약을 삭제하고 처음 상태로 되돌릴까요?", isPresented: $confirmReset, titleVisibility: .visible) {
                Button("초기화", role: .destructive) { store.resetAll() }
            }
        }
    }

    private var profileHeader: some View {
        Section("내 정보") {
            HStack(spacing: 14) {
                ZStack {
                    Circle().fill(store.theme.accent.opacity(0.15)).frame(width: 52, height: 52)
                    Text(String(store.settings.ownerName.prefix(1))).font(.app(.title3, .bold)).foregroundStyle(store.theme.accent)
                }
                VStack(alignment: .leading, spacing: 3) {
                    TextField("이름", text: Binding(get: { store.settings.ownerName }, set: { v in store.update { $0.ownerName = v } }))
                        .font(.app(.headline, .semibold))
                    Text("\(store.brand.bank) 고객 · 시뮬레이션 계정").font(.app(.caption)).foregroundStyle(.secondary)
                }
            }
            .padding(.vertical, 4)
        }
    }

    private func nameField(_ title: String, symbol: String, keyPath: WritableKeyPath<BankSettings, String>, placeholder: String) -> some View {
        HStack {
            Label(title, systemImage: symbol)
            TextField(placeholder, text: Binding(get: { store.settings[keyPath: keyPath] }, set: { v in store.update { $0[keyPath: keyPath] = v } }))
                .multilineTextAlignment(.trailing)
                .autocorrectionDisabled()
        }
    }

    static var version: String {
        let v = Bundle.main.infoDictionary?["CFBundleShortVersionString"] as? String ?? "1.0"
        let b = Bundle.main.infoDictionary?["CFBundleVersion"] as? String ?? "1"
        return "\(v) (\(b))"
    }
}

struct ThemePicker: View {
    @EnvironmentObject var store: BankStore
    var body: some View {
        List {
            Section {
                ForEach(Theme.all) { t in
                    Button { store.update { $0.themeID = t.id } } label: {
                        HStack(spacing: 12) {
                            RoundedRectangle(cornerRadius: 8)
                                .fill(LinearGradient(colors: t.cardGradient, startPoint: .topLeading, endPoint: .bottomTrailing))
                                .frame(width: 48, height: 32)
                            VStack(alignment: .leading) {
                                Text(t.displayName).font(.app(.body, .semibold))
                                Text(t.style == .toss ? "\(t.bankName) · 다크 리스트 레이아웃" : t.tagline).font(.app(.caption)).foregroundStyle(.secondary)
                            }
                            Spacer()
                            if store.settings.themeID == t.id { Image(systemName: "checkmark").foregroundStyle(store.theme.accent) }
                        }
                    }.buttonStyle(.plain)
                }
            } footer: { Text("모두 이 앱을 위해 만든 가상의 브랜드입니다. 은행 이름은 설정 > 금융사 이름에서 바꿀 수 있습니다.") }
        }
        .navigationTitle("테마")
    }
}

struct TimeFlowSettings: View {
    @EnvironmentObject var store: BankStore
    var body: some View {
        Form {
            Section {
                Picker("속도", selection: Binding(get: { TimeSpeed.allCases.first { $0.rawValue == store.clock.speed } ?? .realtime },
                                               set: { store.setSpeed($0) })) {
                    ForEach(TimeSpeed.allCases) { Text($0.title).tag($0) }
                }
                HStack {
                    Text("시뮬레이션 시각"); Spacer()
                    TimelineView(.periodic(from: .now, by: 1)) { _ in
                        Text(store.now.formatted(date: .abbreviated, time: .shortened)).foregroundStyle(.secondary).monospacedDigit()
                    }
                }
                if store.clock.speed != 1 {
                    Button("실시간 속도로 (시뮬레이션 시각 유지)") { store.resetToRealSpeed() }
                }
            } footer: {
                Text("속도를 올리면 월급·생활비·이자·상환·예약 입금이 그 배속으로 진행됩니다. 1,200x면 실제 한 달 동안 100년이 지납니다. 가장 느린 설정이 실시간이며, 시간은 앞으로만 흐릅니다(점프 없음). 가속 중에는 잠금화면 알림 대신 앱 내 배너로 표시됩니다.")
            }
        }
        .navigationTitle("시간 흐름")
    }
}

struct InfoPage: View {
    let title: String
    let items: [(String, String)]

    var body: some View {
        List {
            ForEach(items, id: \.0) { item in
                VStack(alignment: .leading, spacing: 6) {
                    Text(item.0).font(.app(.body, .semibold))
                    Text(item.1).font(.app(.subheadline)).foregroundStyle(.secondary)
                }
                .padding(.vertical, 4)
            }
        }
        .navigationTitle(title)
    }

    static let notices: [(String, String)] = [
        ("[안내] 이 앱은 뱅킹 시뮬레이터입니다", "표시되는 모든 계좌·잔액·거래·이자·알림은 사용자가 설정한 가상의 데이터이며 실제 금융기관·결제 서비스와 연결되어 있지 않습니다."),
        ("[업데이트] 예약 입금 알림", "예약한 시각에 잠금화면 알림이 오고 잔액이 반영됩니다. 앱이 꺼져 있던 기간의 예약도 다음 실행 때 자동 정산됩니다."),
        ("[업데이트] 재력 설정", "설정 > 시뮬레이션 > 재력 설정에서 거지~재벌 중 하나를 고르면 그 생활에 맞는 입출금이 자동으로 생깁니다."),
    ]
    static let faq: [(String, String)] = [
        ("잔액이 실제 돈인가요?", "아닙니다. 모든 금액은 기기 안에서만 계산되는 모의 데이터입니다."),
        ("알림이 오지 않아요", "설정 > 보안·알림에서 입금 알림을 허용했는지, iOS 설정의 알림 권한이 켜져 있는지 확인하세요. 시간 가속 중에는 잠금화면 알림 대신 앱 내 배너로 표시됩니다."),
        ("은행 이름을 바꾸고 싶어요", "설정 > 금융사 이름에서 은행·증권사·카드사·캐피탈 이름을 직접 입력할 수 있습니다."),
        ("데이터를 처음부터 다시 시작하려면?", "설정 > 시뮬레이션 > 모든 데이터 초기화를 누르세요."),
    ]
    static let terms: [(String, String)] = [
        ("이용안내", "이 앱은 금융 서비스가 아니며 실제 자금의 이동이나 금융 거래가 발생하지 않습니다. 화면의 SIM 배지는 시뮬레이션 표시입니다."),
        ("개인정보", "입력한 이름·계좌·거래 데이터는 이 기기 안에만 저장되며 외부로 전송되지 않습니다."),
        ("브랜드", "Nova Bank · Heritage Trust · Orbit Pay 및 기본 금융사 이름은 가상의 표기입니다."),
    ]
}

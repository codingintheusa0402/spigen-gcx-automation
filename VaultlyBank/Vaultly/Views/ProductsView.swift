import SwiftUI

struct ProductsView: View {
    @EnvironmentObject var store: BankStore
    @State private var opening: AccountKind?

    var body: some View {
        NavigationStack {
            ScrollView {
                VStack(spacing: 20) {
                    Text("가입하면 매일 이자가 쌓이고, 매월 자동이체·상환이 실제로 잔액에 반영됩니다.")
                        .font(.footnote).foregroundStyle(.secondary).frame(maxWidth: .infinity, alignment: .leading)
                    ForEach([AccountKind.installmentSavings, .deposit, .loan]) { kind in
                        Button { opening = kind } label: { ProductCard(kind: kind) }.buttonStyle(.plain)
                    }
                    if !store.productAccounts.isEmpty {
                        VStack(alignment: .leading, spacing: 12) {
                            Text("가입 중인 상품").font(.headline)
                            ForEach(store.productAccounts) { acc in
                                NavigationLink(value: acc.id) { AccountTile(account: acc) }.buttonStyle(.plain)
                            }
                        }
                    }
                }.padding()
            }
            .background(store.theme.background)
            .navigationTitle("금융 상품")
            .navigationDestination(for: UUID.self) { AccountDetailView(accountID: $0) }
            .toolbar { ToolbarItem(placement: .topBarLeading) { SimBadge() } }
            .sheet(item: $opening) { ProductEditor(kind: $0) }
        }
    }
}

struct ProductCard: View {
    @EnvironmentObject var store: BankStore
    let kind: AccountKind
    var body: some View {
        Card {
            HStack(spacing: 14) {
                Image(systemName: kind.symbol).font(.title).foregroundStyle(store.theme.accent).frame(width: 44)
                VStack(alignment: .leading, spacing: 4) {
                    Text(title).font(.headline)
                    Text(desc).font(.caption).foregroundStyle(.secondary)
                }
                Spacer()
                Image(systemName: "chevron.right").foregroundStyle(.tertiary)
            }
        }
    }
    private var title: String {
        switch kind {
        case .installmentSavings: return "Vaultly 자유적금"
        case .deposit: return "Vaultly 정기예금"
        case .loan: return "Vaultly 신용대출"
        default: return ""
        }
    }
    private var desc: String {
        switch kind {
        case .installmentSavings: return "매월 자동이체 · 연 최대 7.0% · 만기 이자 지급"
        case .deposit: return "목돈 거치 · 연 최대 5.0% · 일 단위 이자 누적"
        case .loan: return "원리금균등상환 · 연 3.5%~ · 매월 자동 상환"
        default: return ""
        }
    }
}

struct ProductEditor: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    let kind: AccountKind

    @State private var name = ""
    @State private var rate = 3.5
    @State private var months = 12
    @State private var amount: Decimal = 0
    @State private var linked: UUID?
    @State private var start = Date()

    var body: some View {
        NavigationStack {
            Form {
                Section("상품") {
                    TextField("상품 이름", text: $name)
                    HStack {
                        Text("금리"); Spacer()
                        Text(String(format: "연 %.2f%%", rate)).monospacedDigit().fontWeight(.semibold)
                    }
                    Slider(value: $rate, in: 0...20, step: 0.1)
                    Stepper("기간: \(months)개월", value: $months, in: 1...360)
                    AmountField(title: amountTitle, amount: $amount)
                    DatePicker("가입일", selection: $start, displayedComponents: .date)
                }
                Section {
                    Picker("연결 계좌", selection: $linked) {
                        Text("없음").tag(UUID?.none)
                        ForEach(store.checkingAccounts) { a in Text(a.name).tag(Optional(a.id)) }
                    }
                } footer: { Text(linkedFooter) }
                Section("예상") { projection }
                Section {
                    Button { open() } label: { Text("가입하기").frame(maxWidth: .infinity).fontWeight(.bold) }
                        .buttonStyle(.borderedProminent).disabled(amount <= 0 || name.isEmpty).listRowInsets(EdgeInsets())
                } footer: { Text("가입일을 과거로 두면 그동안의 이자·납입·상환이 즉시 정산됩니다.") }
            }
            .navigationTitle(kind.title + " 가입")
            .toolbar { ToolbarItem(placement: .cancellationAction) { Button("취소") { dismiss() } } }
            .onAppear {
                start = store.now
                linked = store.checkingAccounts.first?.id
                switch kind {
                case .installmentSavings: name = "Vaultly 자유적금"; rate = 5.0; amount = 500_000
                case .deposit: name = "Vaultly 정기예금"; rate = 3.8; amount = 10_000_000
                case .loan: name = "Vaultly 신용대출"; rate = 4.2; months = 36; amount = 30_000_000
                default: break
                }
            }
        }
    }

    private var amountTitle: String {
        switch kind { case .installmentSavings: return "월 납입액"; case .deposit: return "예치 금액"; default: return "대출 금액" }
    }
    private var linkedFooter: String {
        switch kind {
        case .installmentSavings: return "매월 가입일에 연결 계좌에서 자동이체되고, 만기 시 원금+이자가 연결 계좌로 돌아옵니다."
        case .deposit: return "예치 금액이 연결 계좌에서 빠져나가고, 만기 시 원금+이자가 돌아옵니다."
        default: return "대출금이 연결 계좌로 입금되고, 매월 원리금이 연결 계좌에서 자동 상환됩니다."
        }
    }

    @ViewBuilder private var projection: some View {
        let terms = ProductTerms(annualRate: rate / 100, termMonths: months, startDate: start, principal: amount,
                                 monthlyAmount: amount, linkedAccountID: nil, lastAccrualDate: start, lastInstallmentDate: start)
        let tax = store.settings.interestTaxRate
        switch kind {
        case .loan:
            let pay = InterestEngine.annuityPayment(principal: amount, annualRate: rate / 100, months: months)
            row("월 상환액", pay)
            row("총 상환액", pay * Decimal(months))
            row("총 이자", pay * Decimal(months) - amount)
        default:
            let p = InterestEngine.projectedMaturity(terms: terms, kind: kind)
            row("만기 원금", p.principal)
            row("만기 이자 (세전)", p.interest)
            row(String(format: "세후 수령액 (세율 %.1f%%)", tax * 100), p.principal + p.interest - (p.interest * Decimal(double: tax)).rounded())
            HStack { Text("만기일").foregroundStyle(.secondary); Spacer(); Text(terms.maturityDate.formatted(date: .long, time: .omitted)) }
        }
    }
    private func row(_ k: String, _ v: Decimal) -> some View {
        HStack { Text(k).foregroundStyle(.secondary); Spacer(); MoneyText(amount: v, font: .body.weight(.semibold), hideable: false) }
    }

    private func open() {
        store.openProduct(kind: kind, name: name, annualRate: rate / 100, termMonths: months, amount: amount,
                          linkedAccountID: linked, startDate: start)
        dismiss()
    }
}

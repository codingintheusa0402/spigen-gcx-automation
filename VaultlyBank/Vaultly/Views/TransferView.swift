import SwiftUI

struct TransferView: View {
    var body: some View { NavigationStack { TransferForm() } }
}

struct TransferForm: View {
    @EnvironmentObject var store: BankStore
    var presetFrom: UUID? = nil

    @State private var from: UUID?
    @State private var toOwn = true
    @State private var toAccount: UUID?
    @State private var recipient = ""
    @State private var bank = "Nova Bank"
    @State private var amount: Decimal = 0
    @State private var memo = ""
    @State private var confirm = false
    @State private var done = false
    @State private var failed = false

    private let banks = ["Nova Bank", "Heritage Trust", "Orbit Pay", "Lumen Credit", "Pioneer Savings"]

    var body: some View {
        Group {
            Form {
                Section("출금 계좌") {
                    Picker("출금 계좌", selection: $from) {
                        ForEach(store.checkingAccounts) { a in
                            Text("\(a.name) · \(Money.format(a.balance, code: store.currency))").tag(Optional(a.id))
                        }
                    }
                }
                Section("받는 사람") {
                    Picker("대상", selection: $toOwn) { Text("내 계좌").tag(true); Text("다른 사람").tag(false) }.pickerStyle(.segmented)
                    if toOwn {
                        Picker("입금 계좌", selection: $toAccount) {
                            ForEach(store.data.accounts.filter { $0.id != from && $0.kind != .loan }) { a in Text(a.name).tag(Optional(a.id)) }
                        }
                    } else {
                        TextField("받는 분 이름", text: $recipient)
                        Picker("은행 (가상)", selection: $bank) { ForEach(banks, id: \.self) { Text($0) } }
                    }
                }
                Section("금액") {
                    AmountField(title: "보낼 금액", amount: $amount)
                    TextField("메모 (선택)", text: $memo)
                }
                Section {
                    Button { confirm = true } label: { Text("보내기").frame(maxWidth: .infinity).fontWeight(.bold) }
                        .buttonStyle(.borderedProminent)
                        .disabled(!canSend)
                        .listRowInsets(EdgeInsets())
                } footer: {
                    Text("모의 송금입니다. 실제 자금 이동은 발생하지 않습니다.")
                }
            }
            .navigationTitle("송금")
            .toolbar { ToolbarItem(placement: .topBarLeading) { SimBadge() } }
            .onAppear {
                if from == nil { from = presetFrom ?? store.checkingAccounts.first?.id }
                if toAccount == nil { toAccount = store.data.accounts.first { $0.id != from && $0.kind != .loan }?.id }
            }
            .confirmationDialog(confirmText, isPresented: $confirm, titleVisibility: .visible) {
                Button("확인") { send() }
            }
            .sheet(isPresented: $done) {
                VStack(spacing: 20) {
                    SuccessCheck()
                    Text("송금 완료").font(.title2.bold())
                    MoneyText(amount: amount, font: .largeTitle.weight(.heavy), hideable: false)
                    Text(toOwn ? (store.account(toAccount ?? UUID())?.name ?? "") : "\(recipient) · \(bank)").foregroundStyle(.secondary)
                    Button("확인") { done = false; amount = 0; memo = "" }.buttonStyle(.borderedProminent).padding(.top)
                }
                .padding(32).presentationDetents([.medium])
            }
            .alert("송금 실패", isPresented: $failed) { Button("확인") {} } message: { Text("잔액이 부족하거나 입력이 올바르지 않습니다.") }
        }
    }

    private var canSend: Bool {
        guard amount > 0, from != nil else { return false }
        return toOwn ? toAccount != nil : !recipient.isEmpty
    }
    private var confirmText: String {
        let target = toOwn ? (store.account(toAccount ?? UUID())?.name ?? "") : "\(recipient) (\(bank))"
        return "\(target)에게 \(Money.format(amount, code: store.currency))을 보낼까요?"
    }
    private func send() {
        guard let from else { return }
        let target: BankStore.TransferTarget = toOwn ? .own(toAccount ?? UUID()) : .external(name: recipient, bank: bank)
        if store.transfer(from: from, to: target, amount: amount, memo: memo) { done = true } else { failed = true }
    }
}

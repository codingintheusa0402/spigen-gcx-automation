import SwiftUI

struct AccountDetailView: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    let accountID: UUID
    @State private var editBalance = false
    @State private var newBalance: Decimal = 0
    @State private var rename = false
    @State private var newName = ""
    @State private var confirmDelete = false

    var body: some View {
        if let account = store.account(accountID) {
            ScrollView {
                VStack(spacing: 20) {
                    AccountTile(account: account)
                    if let terms = account.terms { productInfo(account, terms) }
                    if account.kind == .checking {
                        HStack(spacing: 12) {
                            Button { newBalance = account.balance; editBalance = true } label: {
                                Label("잔액 설정", systemImage: "slider.horizontal.3").frame(maxWidth: .infinity)
                            }.buttonStyle(.borderedProminent)
                            NavigationLink { TransferForm(presetFrom: account.id) } label: {
                                Label("송금", systemImage: "paperplane").frame(maxWidth: .infinity)
                            }.buttonStyle(.bordered)
                        }
                    }
                    Card {
                        Text("거래 내역").font(.headline)
                        let txs = store.transactions(for: accountID)
                        if txs.isEmpty { Text("거래 내역이 없습니다").foregroundStyle(.secondary).padding(.top, 4) }
                        ForEach(txs) { tx in
                            TransactionRow(tx: tx)
                            if tx.id != txs.last?.id { Divider() }
                        }
                    }
                }
                .padding()
            }
            .background(store.theme.background)
            .navigationTitle(account.name)
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .topBarTrailing) {
                    Menu {
                        Button { newName = account.name; rename = true } label: { Label("이름 변경", systemImage: "pencil") }
                        Button(role: .destructive) { confirmDelete = true } label: { Label("계좌 삭제", systemImage: "trash") }
                    } label: { Image(systemName: "ellipsis.circle") }
                }
            }
            .alert("잔액 설정", isPresented: $editBalance) {
                TextField("금액", value: $newBalance, format: .number).keyboardType(.decimalPad)
                Button("저장") { store.setBalance(newBalance, of: accountID) }
                Button("취소", role: .cancel) {}
            } message: { Text("원하는 잔액을 입력하면 차액이 거래 내역에 '잔액 조정'으로 남습니다.") }
            .alert("이름 변경", isPresented: $rename) {
                TextField("계좌 이름", text: $newName)
                Button("저장") { store.rename(accountID, to: newName) }
                Button("취소", role: .cancel) {}
            }
            .confirmationDialog("이 계좌와 모든 거래 내역을 삭제할까요?", isPresented: $confirmDelete, titleVisibility: .visible) {
                Button("삭제", role: .destructive) { dismiss(); store.deleteAccount(accountID) }
            }
        } else {
            ContentUnavailableView("계좌를 찾을 수 없습니다", systemImage: "questionmark.folder")
        }
    }

    @ViewBuilder
    private func productInfo(_ account: Account, _ terms: ProductTerms) -> some View {
        let projection = InterestEngine.projectedMaturity(terms: terms, kind: account.kind)
        let tax = store.settings.interestTaxRate
        Card {
            VStack(alignment: .leading, spacing: 10) {
                Text("상품 정보").font(.headline)
                row("금리", String(format: "연 %.2f%%", terms.annualRate * 100))
                row("기간", "\(terms.termMonths)개월")
                row("가입일", terms.startDate.formatted(date: .long, time: .omitted))
                row("만기일", terms.maturityDate.formatted(date: .long, time: .omitted))
                if let linked = terms.linkedAccountID, let l = store.account(linked) { row("연결 계좌", l.name) }
                Divider()
                switch account.kind {
                case .loan:
                    moneyRow("대출 원금", terms.principal)
                    moneyRow("월 상환액 (원리금균등)", terms.monthlyAmount)
                    moneyRow("남은 원금", account.balance)
                    if let next = Calendar.current.date(byAdding: .month, value: 1, to: terms.lastInstallmentDate), !terms.isClosed {
                        row("다음 상환일", next.formatted(date: .abbreviated, time: .omitted))
                    }
                case .installmentSavings, .deposit:
                    if account.kind == .installmentSavings { moneyRow("월 납입액", terms.monthlyAmount) }
                    moneyRow("현재까지 누적 이자", terms.accruedInterest.rounded())
                    moneyRow("만기 예상 이자 (세전)", projection.interest)
                    moneyRow("만기 예상 수령액 (세후)", projection.principal + projection.interest - (projection.interest * Decimal(double: tax)).rounded())
                    if terms.isClosed { Text("만기 처리 완료 — 이자가 지급되었습니다.").font(.caption).foregroundStyle(.green) }
                default: EmptyView()
                }
            }
        }
    }

    private func row(_ k: String, _ v: String) -> some View {
        HStack { Text(k).foregroundStyle(.secondary); Spacer(); Text(v).fontWeight(.medium) }.font(.subheadline)
    }
    private func moneyRow(_ k: String, _ v: Decimal) -> some View {
        HStack { Text(k).foregroundStyle(.secondary); Spacer(); MoneyText(amount: v, font: .subheadline.weight(.medium)) }.font(.subheadline)
    }
}

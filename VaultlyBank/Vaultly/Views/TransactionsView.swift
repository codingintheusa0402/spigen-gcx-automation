import SwiftUI

enum HistoryPeriod: Int, CaseIterable, Identifiable {
    case d1 = 1, d7 = 7, d30 = 30, m3 = 90, m6 = 180, m12 = 365
    var id: Int { rawValue }
    var title: String {
        switch self {
        case .d1: return "1일"
        case .d7: return "7일"
        case .d30: return "30일"
        case .m3: return "3개월"
        case .m6: return "6개월"
        case .m12: return "12개월"
        }
    }
}

/// 전체 거래내역: period chips (1일 … 12개월) and a category filter row, grouped by day.
struct TransactionsView: View {
    @EnvironmentObject var store: BankStore
    var accountID: UUID? = nil
    @State private var period: HistoryPeriod = .d30
    @State private var category: SpendingCategory? = nil
    @State private var kind: Kind = .all

    enum Kind: String, CaseIterable, Identifiable { case all = "전체", out = "출금", `in` = "입금"; var id: String { rawValue } }

    private var base: [Transaction] {
        let from = Calendar.current.date(byAdding: .day, value: -period.rawValue, to: store.now) ?? store.now
        return store.data.transactions.filter { ($0.date >= from) && (accountID == nil || $0.accountID == accountID) }.sorted { $0.date > $1.date }
    }
    private var filtered: [Transaction] {
        base.filter { tx in
            (kind == .all || (kind == .out ? tx.amount < 0 : tx.amount > 0)) &&
            (category == nil || SpendingReport.category(of: tx) == category)
        }
    }
    private var categoriesPresent: [SpendingCategory] {
        var seen: [SpendingCategory: Int] = [:]
        for tx in base { seen[SpendingReport.category(of: tx), default: 0] += 1 }
        return SpendingCategory.allCases.filter { seen[$0] != nil }
    }

    var body: some View {
        ScrollView {
            VStack(spacing: 12) {
                Picker("기간", selection: $period) { ForEach(HistoryPeriod.allCases) { Text($0.title).tag($0) } }.pickerStyle(.segmented)
                HStack {
                    Picker("종류", selection: $kind) { ForEach(Kind.allCases) { Text($0.rawValue).tag($0) } }.pickerStyle(.segmented).frame(width: 180)
                    Spacer()
                    summary
                }
                ScrollView(.horizontal, showsIndicators: false) {
                    HStack(spacing: 8) {
                        chip(nil, "전체")
                        ForEach(categoriesPresent) { c in chip(c, c.title) }
                    }
                }
                Card {
                    if store.theme.style == .toss {
                        TossTransactionList(transactions: filtered, showAccount: accountID == nil, tappable: true)
                    } else {
                        if filtered.isEmpty { Text("거래 내역이 없습니다").foregroundStyle(.secondary) }
                        ForEach(filtered) { tx in
                            NavigationLink(value: tx) { TransactionRow(tx: tx, showAccount: accountID == nil) }.buttonStyle(.plain)
                            if tx.id != filtered.last?.id { Divider() }
                        }
                    }
                }
            }
            .padding()
        }
        .background(store.theme.background)
        .navigationTitle(accountID == nil ? "전체 거래내역" : "거래내역")
        .navigationBarTitleDisplayMode(.inline)
    }

    private var summary: some View {
        let out = filtered.filter { $0.amount < 0 }.reduce(Decimal(0)) { $0 - $1.amount }
        let inn = filtered.filter { $0.amount > 0 }.reduce(Decimal(0)) { $0 + $1.amount }
        return VStack(alignment: .trailing, spacing: 1) {
            Text("출금 \(Money.format(out, code: store.currency))").font(.app(.caption2)).foregroundStyle(.secondary)
            Text("입금 \(Money.format(inn, code: store.currency))").font(.app(.caption2)).foregroundStyle(store.theme.accent)
        }
    }

    private func chip(_ c: SpendingCategory?, _ title: String) -> some View {
        let on = category == c
        return Button { category = c } label: {
            HStack(spacing: 4) {
                if let c { Image(systemName: c.symbol).font(.app(.caption2)) }
                Text(title).font(.app(.caption, .semibold))
            }
            .padding(.horizontal, 12).padding(.vertical, 7)
            .background(on ? store.theme.accent : Color.primary.opacity(0.07), in: Capsule())
            .foregroundStyle(on ? store.theme.onAccent : .primary)
        }.buttonStyle(.plain)
    }
}

/// 거래 상세 — what a real bank shows when you tap a row.
struct TransactionDetailView: View {
    @EnvironmentObject var store: BankStore
    let tx: Transaction

    var body: some View {
        let cat = SpendingReport.category(of: tx)
        ScrollView {
            VStack(spacing: 16) {
                VStack(spacing: 6) {
                    ZStack {
                        Circle().fill(store.theme.accent.opacity(0.12)).frame(width: 56, height: 56)
                        Image(systemName: cat.symbol).font(.app(.title3)).foregroundStyle(store.theme.accent)
                    }
                    Text(tx.counterparty).font(.app(.title3, .bold)).multilineTextAlignment(.center)
                    MoneyText(amount: tx.amount, signed: true, font: .app(size: 32, .semibold), hideable: false)
                        .foregroundStyle(tx.amount > 0 ? store.theme.accent : .primary)
                    Text(tx.date.formatted(date: .long, time: .shortened)).font(.app(.subheadline)).foregroundStyle(.secondary)
                }
                .padding(.top, 12)
                Card {
                    VStack(spacing: 0) {
                        row("거래 유형", tx.category.title)
                        row("분류", cat.title)
                        row("계좌", store.account(tx.accountID)?.name ?? "-")
                        row("계좌번호", store.account(tx.accountID)?.number ?? "-")
                        if !tx.memo.isEmpty { row("메모", tx.memo) }
                        row("거래 후 잔액", Money.format(tx.balanceAfter, code: store.currency))
                        row("거래 번호", String(tx.id.uuidString.prefix(8)).uppercased(), last: true)
                    }
                }
                Text("모든 거래는 시뮬레이션이며 실제 자금 이동이 없습니다.").font(.app(.caption2)).foregroundStyle(.tertiary)
            }
            .padding()
        }
        .background(store.theme.background)
        .navigationTitle("거래 상세")
        .navigationBarTitleDisplayMode(.inline)
    }

    private func row(_ k: String, _ v: String, last: Bool = false) -> some View {
        VStack(spacing: 0) {
            HStack(alignment: .top) {
                Text(k).font(.app(.subheadline)).foregroundStyle(.secondary)
                Spacer()
                Text(v).font(.app(.subheadline, .semibold)).multilineTextAlignment(.trailing)
            }
            .padding(.vertical, 10)
            if !last { Divider() }
        }
    }
}

import SwiftUI

/// Flat list row: icon · bold amount / gray name · optional pill, as in Korean neobank apps.
struct TossAccountRow: View {
    @EnvironmentObject var store: BankStore
    let account: Account

    var body: some View {
        HStack(spacing: 14) {
            ZStack {
                Circle().fill(store.theme.gradient(for: account.colorIndex)).frame(width: 40, height: 40)
                Image(systemName: account.kind.symbol).font(.app(size: 16, .semibold)).foregroundStyle(.white)
            }
            VStack(alignment: .leading, spacing: 3) {
                MoneyText(amount: account.kind == .loan ? -account.balance : account.balance, font: .app(.title3, .bold))
                HStack(spacing: 4) {
                    Text(account.name).font(.app(.subheadline)).foregroundStyle(.secondary).lineLimit(1)
                    if let t = account.terms, t.accruedInterest > 0, !t.isClosed {
                        Text("+이자 \(Money.format(t.accruedInterest.rounded(), code: store.currency))")
                            .font(.app(.subheadline)).foregroundStyle(store.theme.accent)
                    }
                }
            }
            Spacer(minLength: 8)
            if account.kind == .checking {
                NavigationLink { TransferForm(presetFrom: account.id) } label: { TossPill("송금") }.buttonStyle(.plain)
            }
        }
        .padding(.vertical, 6)
    }
}

struct TossPill: View {
    let title: String
    var prominent = false
    init(_ title: String, prominent: Bool = false) { self.title = title; self.prominent = prominent }
    var body: some View {
        Text(title)
            .font(.app(.subheadline, .semibold))
            .padding(.horizontal, 16).padding(.vertical, 9)
            .background(prominent ? Color.accentColor.opacity(0.18) : Color.primary.opacity(0.08), in: RoundedRectangle(cornerRadius: 10, style: .continuous))
            .foregroundStyle(prominent ? Color.accentColor : Color.primary)
    }
}

/// Transactions grouped by day: date in the left gutter, name + time, amount + balance after.
struct TossTransactionList: View {
    @EnvironmentObject var store: BankStore
    let transactions: [Transaction]
    var showAccount = false
    var tappable = true

    private var days: [(String, [Transaction])] {
        let cal = Calendar.current
        var order: [String] = []
        var groups: [String: [Transaction]] = [:]
        for tx in transactions.sorted(by: { $0.date > $1.date }) {
            let key = String(format: "%02d.%02d", cal.component(.month, from: tx.date), cal.component(.day, from: tx.date))
            if groups[key] == nil { order.append(key) }
            groups[key, default: []].append(tx)
        }
        return order.map { ($0, groups[$0] ?? []) }
    }

    var body: some View {
        if transactions.isEmpty {
            Text("거래 내역이 없습니다").foregroundStyle(.secondary).padding(.vertical, 8)
        }
        VStack(spacing: 0) {
            ForEach(days, id: \.0) { day, txs in
                ForEach(Array(txs.enumerated()), id: \.element.id) { idx, tx in
                    NavigationLink(value: tx) {
                    HStack(alignment: .top, spacing: 12) {
                        Text(idx == 0 ? day : "").font(.app(.subheadline)).foregroundStyle(.secondary).frame(width: 44, alignment: .leading)
                        VStack(alignment: .leading, spacing: 3) {
                            Text(tx.counterparty).font(.app(.body, .semibold)).lineLimit(1)
                            Text(subtitle(tx)).font(.app(.subheadline)).foregroundStyle(.secondary).lineLimit(1)
                        }
                        Spacer(minLength: 8)
                        VStack(alignment: .trailing, spacing: 3) {
                            MoneyText(amount: tx.amount, signed: false, font: .app(.body, .semibold))
                                .foregroundStyle(tx.amount > 0 ? store.theme.accent : (tx.amount == 0 ? Color.secondary : Color.primary))
                            MoneyText(amount: tx.balanceAfter, font: .app(.subheadline)).foregroundStyle(.secondary)
                        }
                    }
                    .padding(.vertical, 10)
                    .contentShape(Rectangle())
                    }
                    .buttonStyle(.plain)
                    .disabled(!tappable)
                }
            }
        }
    }

    private func subtitle(_ tx: Transaction) -> String {
        var parts = [tx.date.formatted(date: .omitted, time: .shortened)]
        if !tx.memo.isEmpty { parts.append(tx.memo) }
        if showAccount, let a = store.account(tx.accountID) { parts.append(a.name) }
        return parts.joined(separator: " · ")
    }
}

/// Header line used instead of the navigation bar on TOSS-style screens.
struct TossHeader<Trailing: View>: View {
    @EnvironmentObject var store: BankStore
    let title: String
    var symbol: String? = nil
    var subtitle: AnyView? = nil
    @ViewBuilder var trailing: () -> Trailing

    var body: some View {
        HStack(spacing: 8) {
            if let symbol {
                ZStack {
                    Circle().fill(store.theme.accent).frame(width: 26, height: 26)
                    Image(systemName: symbol).font(.app(size: 12, .bold)).foregroundStyle(.white)
                }
            }
            Text(title).font(.app(.title3, .bold))
            if let subtitle { subtitle }
            Spacer()
            trailing()
        }
        .padding(.top, 6)
    }
}

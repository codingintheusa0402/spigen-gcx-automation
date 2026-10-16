import SwiftUI

struct HomeView: View {
    @EnvironmentObject var store: BankStore
    @State private var showAddAccount = false

    var body: some View {
        if store.theme.style == .toss { tossBody } else { cardBody }
    }

    private var tossBody: some View {
        NavigationStack {
            ScrollView {
                VStack(spacing: 14) {
                    TossHeader(title: store.brand.bank, symbol: "building.columns.fill",
                               subtitle: AnyView(Image(systemName: "chevron.right").font(.app(.caption, .semibold)).foregroundStyle(.secondary))) {
                        Button { store.update { $0.hideBalances.toggle() } } label: {
                            Image(systemName: store.settings.hideBalances ? "eye.slash" : "eye").font(.app(.title3)).foregroundStyle(.primary)
                        }
                        NavigationLink { TransferForm() } label: { Image(systemName: "paperplane").font(.app(.title3)).foregroundStyle(.primary) }
                    }
                    Card {
                        VStack(alignment: .leading, spacing: 10) {
                            HStack(alignment: .firstTextBaseline) {
                                Text("총자산").font(.app(.subheadline)).foregroundStyle(.secondary)
                                Spacer()
                            }
                            MoneyText(amount: store.totalAssets, font: .app(size: 30, .bold))
                            if store.totalDebt > 0 {
                                HStack(spacing: 4) {
                                    Text("대출").font(.app(.caption)).foregroundStyle(.secondary)
                                    MoneyText(amount: store.totalDebt, font: .app(.caption, .semibold)).foregroundStyle(.secondary)
                                }
                            }
                        }
                    }
                    Card {
                        VStack(spacing: 4) {
                            ForEach(store.checkingAccounts + store.productAccounts) { acc in
                                NavigationLink(value: acc.id) { TossAccountRow(account: acc) }.buttonStyle(.plain)
                            }
                            Divider().padding(.vertical, 6)
                            HStack {
                                Text("계좌 \(store.checkingAccounts.count)").font(.app(.subheadline)).foregroundStyle(.secondary)
                                Text("|").foregroundStyle(.quaternary)
                                Text("상품 \(store.productAccounts.count)").font(.app(.subheadline)).foregroundStyle(.secondary)
                                Spacer()
                                Button { showAddAccount = true } label: {
                                    Text("계좌 추가").font(.app(.subheadline)).foregroundStyle(.secondary)
                                }
                            }
                        }
                    }
                    upcoming
                    Card {
                        VStack(alignment: .leading, spacing: 6) {
                            Text("최근 거래").font(.app(.title3, .bold))
                            TossTransactionList(transactions: Array(store.recentTransactions.prefix(10)), showAccount: true)
                        }
                    }
                }
                .padding(.horizontal).padding(.bottom)
            }
            .background(store.theme.background)
            .toolbar(.hidden, for: .navigationBar)
            .navigationDestination(for: UUID.self) { AccountDetailView(accountID: $0) }
            .sheet(isPresented: $showAddAccount) { AddAccountSheet() }
        }
    }

    private var cardBody: some View {
        NavigationStack {
            ScrollView {
                VStack(spacing: 20) {
                    totalCard
                    section("내 계좌", trailing: {
                        HStack(spacing: 14) {
                            NavigationLink { TransferForm() } label: { Image(systemName: "paperplane.fill") }
                            Button { showAddAccount = true } label: { Image(systemName: "plus.circle.fill") }
                        }
                    }) {
                        ForEach(store.checkingAccounts) { acc in
                            NavigationLink(value: acc.id) { AccountTile(account: acc) }.buttonStyle(.plain)
                        }
                    }
                    if !store.productAccounts.isEmpty {
                        section("적금 · 예금 · 대출") {
                            ForEach(store.productAccounts) { acc in
                                NavigationLink(value: acc.id) { AccountTile(account: acc) }.buttonStyle(.plain)
                            }
                        }
                    }
                    upcoming
                    section("최근 거래") {
                        Card {
                            if store.recentTransactions.isEmpty {
                                Text("거래 내역이 없습니다").foregroundStyle(.secondary)
                            }
                            ForEach(store.recentTransactions.prefix(8)) { tx in
                                TransactionRow(tx: tx, showAccount: true)
                                if tx.id != store.recentTransactions.prefix(8).last?.id { Divider() }
                            }
                        }
                    }
                }
                .padding()
            }
            .background(store.theme.background)
            .navigationTitle(store.brand.bank)
            .navigationDestination(for: UUID.self) { AccountDetailView(accountID: $0) }
            .toolbar {
                ToolbarItem(placement: .topBarTrailing) {
                    Button { store.update { $0.hideBalances.toggle() } } label: {
                        Image(systemName: store.settings.hideBalances ? "eye.slash" : "eye")
                    }
                }
            }
            .sheet(isPresented: $showAddAccount) { AddAccountSheet() }
        }
    }

    private var totalCard: some View {
        let on = store.theme.onAccent
        return VStack(alignment: .leading, spacing: 8) {
            HStack {
                Text("\(store.settings.ownerName)님의 총자산").font(.app(.subheadline)).foregroundStyle(on.opacity(0.8))
                Spacer()
                Text(store.theme.tagline).font(.app(.caption2)).foregroundStyle(on.opacity(0.6))
            }
            MoneyText(amount: store.totalAssets, font: .app(size: 32, .bold))
                .foregroundStyle(on)
            HStack(spacing: 4) {
                Text("은행").font(.app(.caption)).foregroundStyle(on.opacity(0.75))
                MoneyText(amount: store.bankAssets, font: .app(.caption, .semibold)).foregroundStyle(on)
                Text("· 투자").font(.app(.caption)).foregroundStyle(on.opacity(0.75))
                MoneyText(amount: store.portfolio.investedKRW + store.portfolio.cashKRWTotal, font: .app(.caption, .semibold)).foregroundStyle(on)
            }
            if store.totalDebt > 0 {
                HStack(spacing: 4) {
                    Text("대출 잔액").font(.app(.caption)).foregroundStyle(on.opacity(0.75))
                    MoneyText(amount: store.totalDebt, font: .app(.caption, .semibold)).foregroundStyle(on)
                }
            }
        }
        .padding(20)
        .frame(maxWidth: .infinity, alignment: .leading)
        .background(store.theme.accent)
        .clipShape(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
    }

    @ViewBuilder private var upcoming: some View {
        let next = store.data.events.filter { $0.isEnabled }.sorted { $0.nextFireDate < $1.nextFireDate }.prefix(3)
        if !next.isEmpty {
            section("다가오는 예약") {
                Card {
                    ForEach(Array(next)) { ev in
                        HStack {
                            Image(systemName: "clock.badge").foregroundStyle(store.theme.accent)
                            VStack(alignment: .leading) {
                                Text("\(ev.counterparty) · \(ev.memo)").font(.app(.subheadline, .semibold))
                                Text(ev.nextFireDate.formatted(date: .abbreviated, time: .shortened)).font(.app(.caption)).foregroundStyle(.secondary)
                            }
                            Spacer()
                            MoneyText(amount: ev.amount, signed: true, font: .app(.subheadline, .bold))
                        }
                        if ev.id != next.last?.id { Divider() }
                    }
                }
            }
        }
    }

    private func section<C: View>(_ title: String, @ViewBuilder content: () -> C) -> some View {
        section(title, trailing: { EmptyView() }, content: content)
    }

    private func section<T: View, C: View>(_ title: String, @ViewBuilder trailing: () -> T, @ViewBuilder content: () -> C) -> some View {
        VStack(alignment: .leading, spacing: 12) {
            HStack { Text(title).font(.app(.headline, .semibold)); Spacer(); trailing() }
            content()
        }
    }
}

struct AddAccountSheet: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    @State private var name = "새 입출금통장"
    @State private var balance: Decimal = 0
    @State private var color = 0

    var body: some View {
        NavigationStack {
            Form {
                TextField("계좌 이름", text: $name)
                AmountField(title: "시작 잔액", amount: $balance)
                HStack(spacing: 12) {
                    Text("카드 색상").foregroundStyle(.secondary)
                    Spacer()
                    ForEach(0..<6, id: \.self) { i in
                        Circle().fill(store.theme.gradient(for: i)).frame(width: 28, height: 28)
                            .overlay(Circle().stroke(Color.primary, lineWidth: color == i ? 2 : 0))
                            .onTapGesture { color = i }
                    }
                }
            }
            .navigationTitle("계좌 추가")
            .toolbar {
                ToolbarItem(placement: .cancellationAction) { Button("취소") { dismiss() } }
                ToolbarItem(placement: .confirmationAction) {
                    Button("추가") { store.addChecking(name: name, balance: balance, colorIndex: color); dismiss() }
                        .disabled(name.isEmpty)
                }
            }
        }
    }
}

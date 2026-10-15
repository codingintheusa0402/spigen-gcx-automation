import SwiftUI

struct HomeView: View {
    @EnvironmentObject var store: BankStore
    @State private var showAddAccount = false

    var body: some View {
        NavigationStack {
            ScrollView {
                VStack(spacing: 20) {
                    totalCard
                    section("내 계좌", trailing: { Button { showAddAccount = true } label: { Image(systemName: "plus.circle.fill") } }) {
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
            .navigationTitle(store.theme.bankName)
            .navigationDestination(for: UUID.self) { AccountDetailView(accountID: $0) }
            .toolbar {
                ToolbarItem(placement: .topBarLeading) { SimBadge() }
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
        VStack(alignment: .leading, spacing: 8) {
            HStack {
                Text("\(store.settings.ownerName)님의 총자산").font(.subheadline).foregroundStyle(.white.opacity(0.8))
                Spacer()
                Text(store.theme.tagline).font(.caption2).foregroundStyle(.white.opacity(0.6))
            }
            MoneyText(amount: store.totalAssets, font: .system(size: 36, weight: .heavy))
                .foregroundStyle(.white)
            if store.totalDebt > 0 {
                HStack(spacing: 4) {
                    Text("대출 잔액").font(.caption).foregroundStyle(.white.opacity(0.75))
                    MoneyText(amount: store.totalDebt, font: .caption.weight(.semibold)).foregroundStyle(.white)
                }
            }
        }
        .padding(22)
        .frame(maxWidth: .infinity, alignment: .leading)
        .background(LinearGradient(colors: store.theme.cardGradient, startPoint: .topLeading, endPoint: .bottomTrailing))
        .clipShape(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
        .shadow(color: store.theme.accent.opacity(0.35), radius: 16, y: 8)
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
                                Text("\(ev.counterparty) · \(ev.memo)").font(.subheadline.weight(.semibold))
                                Text(ev.nextFireDate.formatted(date: .abbreviated, time: .shortened)).font(.caption).foregroundStyle(.secondary)
                            }
                            Spacer()
                            MoneyText(amount: ev.amount, signed: true, font: .subheadline.weight(.bold))
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
            HStack { Text(title).font(.headline); Spacer(); trailing() }
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

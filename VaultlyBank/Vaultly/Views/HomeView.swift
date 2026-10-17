import SwiftUI

/// Value-based routes for the Home stack. Mixing view-based NavigationLinks with
/// navigationDestination(for:) makes the view links re-push when the path changes, so everything is a value.
enum HomeRoute: Hashable {
    case history(UUID?), life, challenge, report, goals, jobs, job(JobKind), schedule
}

extension View {
    func homeRoutes() -> some View {
        self.navigationDestination(for: HomeRoute.self) { r in
            switch r {
            case .history(let id): TransactionsView(accountID: id)
            case .life: LifeView()
            case .challenge: ChallengeView()
            case .report: ReportView()
            case .goals: GoalsView()
            case .jobs: JobsHubView()
            case .job(let kind): JobSessionView(kind: kind)
            case .schedule: DayScheduleView()
            }
        }
        .navigationDestination(for: UUID.self) { AccountDetailView(accountID: $0) }
        .navigationDestination(for: Transaction.self) { TransactionDetailView(tx: $0) }
    }
}

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
                                MockChip()
                                Spacer()
                            }
                            MoneyText(amount: store.totalAssets, font: .app(size: 30, .semibold))
                            if store.totalDebt > 0 {
                                HStack(spacing: 4) {
                                    Text("대출").font(.app(.caption)).foregroundStyle(.secondary)
                                    MoneyText(amount: store.totalDebt, font: .app(.caption, .semibold)).foregroundStyle(.secondary)
                                }
                            }
                        }
                    }
                    lifeRow
                    CharacterCard()
                    if let q = store.pendingQuest { QuestCard(quest: q) }
                    quickTiles
                    goalsCard
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
                            HStack {
                                Text("최근 거래").font(.app(.title3, .bold))
                                Spacer()
                                NavigationLink(value: HomeRoute.history(nil)) { Text("전체보기").font(.app(.subheadline)).foregroundStyle(.secondary) }
                            }
                            TossTransactionList(transactions: Array(store.recentTransactions.prefix(10)), showAccount: true)
                        }
                    }
                }
                .padding(.horizontal).padding(.bottom)
            }
            .background(store.theme.background)
            .toolbar(.hidden, for: .navigationBar)
            .homeRoutes()
            .sheet(isPresented: $showAddAccount) { AddAccountSheet() }
            .sheet(item: $store.questResult) { OutcomeSheet(result: $0).presentationDetents([.medium]) }
            .sheet(isPresented: Binding(get: { store.game.isEnded && !endingShown }, set: { _ in endingShown = true })) { EndingView() }
        }
    }

    @State private var endingShown = false

    /// "만 34세 · 🏠 중산층 · 부자까지 12%" → 내 인생
    @ViewBuilder private var lifeRow: some View {
        if store.lifestyle.tier != nil {
            NavigationLink(value: HomeRoute.life) {
                HStack(spacing: 10) {
                    Text(store.currentTier.emoji)
                    Text("만 \(store.age)세 · \(store.currentTier.title)").font(.app(.subheadline, .semibold))
                    if let next = GameState.nextTier(after: store.currentTier) {
                        let th = GameState.threshold(for: next)
                        let p = th > 0 ? min(1, max(0, (store.netWorth / th).doubleValue)) : 0
                        Text("· \(next.title)까지 \(Int(p * 100))%").font(.app(.caption)).foregroundStyle(.secondary)
                    }
                    Spacer()
                    if store.pendingQuest != nil { Circle().fill(store.theme.accentSecondary).frame(width: 8, height: 8) }
                    Image(systemName: "chevron.right").font(.app(.caption)).foregroundStyle(.tertiary)
                }
                .padding(.horizontal, 14).padding(.vertical, 11)
                .background(store.theme.card, in: RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
                .overlay(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous).stroke(Color.primary.opacity(0.06)))
            }
            .buttonStyle(.plain)
            .accessibilityIdentifier("life-row")
        }
    }

    private var cardBody: some View {
        NavigationStack {
            ScrollView {
                VStack(spacing: 20) {
                    totalCard
                    lifeRow
                    CharacterCard()
                    if let q = store.pendingQuest { QuestCard(quest: q) }
                    quickTiles
                    goalsCard
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
                    section("최근 거래", trailing: { NavigationLink(value: HomeRoute.history(nil)) { Text("전체보기").font(.app(.subheadline)) } }) {
                        Card {
                            if store.recentTransactions.isEmpty {
                                Text("거래 내역이 없습니다").foregroundStyle(.secondary)
                            }
                            ForEach(store.recentTransactions.prefix(8)) { tx in
                                NavigationLink(value: tx) { TransactionRow(tx: tx, showAccount: true) }.buttonStyle(.plain)
                                if tx.id != store.recentTransactions.prefix(8).last?.id { Divider() }
                            }
                        }
                    }
                }
                .padding()
            }
            .background(store.theme.background)
            .navigationTitle(store.brand.bank)
            .homeRoutes()
            .sheet(item: $store.questResult) { OutcomeSheet(result: $0).presentationDetents([.medium]) }
            .sheet(isPresented: Binding(get: { store.game.isEnded && !endingShown }, set: { _ in endingShown = true })) { EndingView() }
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
                MockChip()
                Spacer()
                Text(store.theme.tagline).font(.app(.caption2)).foregroundStyle(on.opacity(0.6))
            }
            MoneyText(amount: store.totalAssets, font: .app(size: 32, .semibold))
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

    /// 무지출 챌린지 + 이번 달 소비 — the two things people open a money app for every day.
    private var quickTiles: some View {
        let now = store.now
        let streak = store.challenge.streak(now: now)
        let saved = store.challenge.total(inMonthOf: now)
        let report = store.report(for: now)
        return VStack(spacing: 12) {
          HStack(spacing: 12) {
            NavigationLink(value: HomeRoute.challenge) {
                quickTile(title: "무지출 챌린지", symbol: "flame.fill", tint: .orange,
                          value: Money.format(saved, code: store.currency),
                          sub: streak > 0 ? "🔥 \(streak)일 연속" : "참은 돈 쌓기")
            }.buttonStyle(.plain).accessibilityIdentifier("tile-challenge")
            NavigationLink(value: HomeRoute.report) {
                quickTile(title: "이번 달 소비", symbol: "chart.pie.fill", tint: store.theme.accent,
                          value: Money.format(report.spent, code: store.currency),
                          sub: report.changeRatio.map { String(format: "지난달 대비 %@%.0f%%", $0 >= 0 ? "+" : "", $0 * 100) } ?? "리포트 보기")
            }.buttonStyle(.plain).accessibilityIdentifier("tile-report")
          }
            NavigationLink(value: HomeRoute.jobs) {
                quickTile(title: "오늘의 알바", symbol: "briefcase.fill", tint: .green,
                          value: store.shiftsLeftToday > 0 ? "\(store.shiftsLeftToday)탕 가능" : "오늘은 끝",
                          sub: "편의점 · 카페 · 스쿠터 · 대리 · 과외 · 단타 · 블랙잭")
            }.buttonStyle(.plain).accessibilityIdentifier("tile-jobs")
        }
    }

    private func quickTile(title: String, symbol: String, tint: Color, value: String, sub: String) -> some View {
        VStack(alignment: .leading, spacing: 6) {
            HStack(spacing: 6) {
                Image(systemName: symbol).font(.app(.caption, .semibold)).foregroundStyle(tint)
                Text(title).font(.app(.caption, .medium)).foregroundStyle(.secondary)
                Spacer()
                Image(systemName: "chevron.right").font(.app(.caption2)).foregroundStyle(.tertiary)
            }
            Text(store.settings.hideBalances ? "••••••" : value).font(.app(.headline, .semibold)).monospacedDigit().lineLimit(1).minimumScaleFactor(0.6)
            Text(sub).font(.app(.caption2)).foregroundStyle(.secondary).lineLimit(1)
        }
        .padding(14)
        .frame(maxWidth: .infinity, alignment: .leading)
        .background(store.theme.card, in: RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
        .overlay(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous).stroke(Color.primary.opacity(0.06)))
    }

    @ViewBuilder private var goalsCard: some View {
        if !store.goals.isEmpty {
            Card {
                VStack(alignment: .leading, spacing: 10) {
                    HStack {
                        Text("저축 목표").font(.app(.headline, .semibold))
                        Spacer()
                        NavigationLink(value: HomeRoute.goals) { Text("전체").font(.app(.subheadline)).foregroundStyle(.secondary) }
                    }
                    ForEach(store.goals.prefix(3)) { g in
                        let bal = store.account(g.accountID)?.balance ?? 0
                        let p = g.progress(balance: bal)
                        VStack(alignment: .leading, spacing: 5) {
                            HStack {
                                Text("\(g.emoji) \(g.title)").font(.app(.subheadline, .semibold))
                                Spacer()
                                Text("\(Int(p * 100))%").font(.app(.caption, .semibold)).foregroundStyle(store.theme.accent)
                            }
                            GeometryReader { geo in
                                Capsule().fill(Color.primary.opacity(0.07))
                                    .overlay(alignment: .leading) { Capsule().fill(store.theme.accent).frame(width: geo.size.width * p) }
                            }.frame(height: 6)
                        }
                    }
                }
            }
        }
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

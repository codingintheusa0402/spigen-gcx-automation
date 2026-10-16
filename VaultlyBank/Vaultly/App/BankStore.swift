import Foundation
import Combine
import SwiftUI

struct BannerMessage: Identifiable, Equatable {
    let id = UUID()
    let title: String
    let body: String
    let positive: Bool
}

/// Single source of truth. Every mutation goes through here so it is persisted
/// and the notification schedule stays in sync.
@MainActor
final class BankStore: ObservableObject {
    @Published private(set) var data: BankData {
        didSet {
            Money.wonSuffix = Theme.named(data.settings.themeID).wonSuffix
            if data != oldValue { Persistence.save(data) }
        }
    }
    @Published var banner: BannerMessage?
    @Published var notificationsAuthorized = false
    @Published var questResult: QuestResult?       // last answered quest → OutcomeSheet (presented by Home/LifeView)

    private var timer: AnyCancellable?

    init() {
        data = Persistence.load() ?? BankData.starter()
        Money.wonSuffix = Theme.named(data.settings.themeID).wonSuffix
        // Check for due events every 20s while the app is in the foreground so the
        // "4:00pm allowance" shows up on screen within seconds of the notification.
        timer = Timer.publish(every: 20, on: .main, in: .common).autoconnect()
            .sink { [weak self] _ in Task { @MainActor in self?.catchUp() } }
        Task { notificationsAuthorized = await NotificationService.authorizationStatus() == .authorized }
    }

    // MARK: Derived

    var settings: BankSettings { data.settings }
    var currency: String { data.settings.currencyCode }
    var theme: Theme { Theme.named(data.settings.themeID) }
    var brand: Brand { data.settings.brand }
    var checkingAccounts: [Account] { data.accounts.filter { $0.kind == .checking } }
    var productAccounts: [Account] { data.accounts.filter { $0.kind.isProduct } }

    var bankAssets: Decimal {
        data.accounts.reduce(0) { $0 + ($1.kind == .loan ? -$1.balance : $1.balance) }
    }
    var totalAssets: Decimal { bankAssets + data.portfolio.investedKRW + data.portfolio.cashKRWTotal }
    var portfolio: Portfolio { data.portfolio }
    var totalDebt: Decimal { data.accounts.filter { $0.kind == .loan }.reduce(0) { $0 + $1.balance } }

    func account(_ id: UUID) -> Account? { data.accounts.first { $0.id == id } }

    func transactions(for id: UUID) -> [Transaction] {
        data.transactions.filter { $0.accountID == id }.sorted { $0.date > $1.date }
    }
    var recentTransactions: [Transaction] { Array(data.transactions.sorted { $0.date > $1.date }.prefix(30)) }

    // MARK: Time-based settlement

    /// Simulated "now" (equals wall-clock time at 1x).
    var now: Date { data.clock.now() }
    var clock: SimClock { data.clock }

    func catchUp() {
        var d = data
        let now = d.clock.now()
        let fired = ScheduleEngine.processDue(now: now, data: &d)
        InterestEngine.accrue(now: now, data: &d, taxRate: d.settings.interestTaxRate)
        MarketSimulator.tick(now: now, portfolio: &d.portfolio)
        let lifeMessages = LifestyleEngine.process(now: now, data: &d)
        // Fast-forward can create tens of thousands of rows; keep the ledger bounded.
        if d.transactions.count > 4000 {
            d.transactions.sort { $0.date < $1.date }
            d.transactions.removeFirst(d.transactions.count - 2500)
        }
        settleGoals(now: now, data: &d)
        if d != data { data = d }
        if let msg = lifeMessages.last {
            banner = BannerMessage(title: msg.hasPrefix("📩") ? "새로운 선택" : (msg.hasPrefix("🏁") ? "인생 결산" : "선택의 결과"),
                                   body: msg.replacingOccurrences(of: "📩 ", with: "").replacingOccurrences(of: "🏁 ", with: ""),
                                   positive: !msg.contains("손실") && !msg.contains("벌금") && !msg.contains("적발"))
        } else if let last = fired.last {
            banner = BannerMessage(title: last.event.isInternal ? "자동이체" : (last.event.isIncoming ? "입금" : "출금"),
                                   body: NotificationService.bankLine(bank: brand.bank, date: last.transaction.date, incoming: last.event.isIncoming && !last.event.isInternal,
                                                                      amount: abs(last.event.amount), counterparty: last.event.counterparty,
                                                                      balance: last.transaction.balanceAfter, currency: currency),
                                   positive: last.event.isIncoming && !last.event.isInternal)
        }
        unlockAchievements(now: now)
        syncNotifications()
        WidgetBridge.publish(store: self)
    }

    private func settleGoals(now: Date, data d: inout BankData) {
        for i in d.goals.indices where d.goals[i].reachedAt == nil {
            if let acc = d.accounts.first(where: { $0.id == d.goals[i].accountID }), acc.balance >= d.goals[i].target {
                d.goals[i].reachedAt = now
            }
        }
    }

    private func unlockAchievements(now: Date) {
        var d = data
        let hits = AchievementEngine.evaluate(now: now, data: &d, totalAssets: totalAssets)
        guard !hits.isEmpty else { return }
        data = d
        if banner == nil, let a = hits.first {
            banner = BannerMessage(title: "🏆 \(a.title)", body: a.detail, positive: true)
        }
    }

    func syncNotifications() {
        // Lock-screen notifications only make sense when simulated time == real time.
        NotificationService.sync(events: data.clock.isRealTime ? data.events : [], accounts: data.accounts, bank: brand.bank, currency: currency)
        LiveActivityService.sync(events: data.clock.isRealTime ? data.events : [], accounts: data.accounts, bank: brand.bank, currency: currency)
        // Quest notifications work at any speed: convert the simulated date to the real instant.
        var upcoming: (Quest, Date)? = nil
        if let pid = data.game.pendingQuestID, let q = QuestLibrary.quest(pid) { upcoming = (q, Date().addingTimeInterval(2)) }
        else if let nid = data.game.nextQuestID, let q = QuestLibrary.quest(nid), let next = data.game.nextQuestDate {
            upcoming = (q, data.clock.realDate(for: next))
        }
        NotificationService.syncQuest(upcoming, pendingAlreadyShown: data.game.pendingQuestID != nil)
    }

    /// Next enabled incoming event (widgets / live activity).
    var nextIncomingEvent: ScheduledEvent? {
        data.events.filter { $0.isEnabled && $0.isIncoming && !$0.isInternal }.min { $0.nextFireDate < $1.nextFireDate }
    }

    // MARK: Time control

    func setSpeed(_ speed: TimeSpeed) {
        var d = data; d.clock.setSpeed(speed.rawValue); data = d
        catchUp()
    }

    /// Back to 1x. Simulated time keeps its current value (it never moves backwards);
    /// from here on it advances at real speed.
    func resetToRealSpeed() {
        var d = data
        d.clock = SimClock(speed: 1, anchorReal: Date(), anchorVirtual: d.clock.now())
        data = d
        syncNotifications()
    }

    func requestNotifications() {
        Task {
            notificationsAuthorized = await NotificationService.requestAuthorization()
            syncNotifications()
        }
    }

    // MARK: Settings

    func update(_ change: (inout BankSettings) -> Void) {
        var s = data.settings; change(&s); data.settings = s
    }

    func resetAll() {
        Persistence.reset()
        data = BankData.starter()
        syncNotifications()
        WidgetBridge.publish(store: self)
    }

    /// Monthly spending report for the month containing `month` (simulated time).
    func report(for month: Date) -> SpendingReport {
        SpendingReport.build(transactions: data.transactions, month: month)
    }

    // MARK: Accounts

    func addChecking(name: String, balance: Decimal, colorIndex: Int) {
        var d = data
        let acc = Account(name: name, kind: .checking, number: Account.generateNumber(), balance: 0, colorIndex: colorIndex)
        d.accounts.append(acc)
        if balance != 0 {
            Ledger.apply(amount: balance, to: acc.id, date: now, counterparty: "직접 설정",
                         memo: "시작 잔액", category: .adjustment, data: &d)
        }
        data = d
    }

    func setBalance(_ balance: Decimal, of id: UUID) {
        var d = data; Ledger.setBalance(balance, of: id, date: now, data: &d); data = d
    }

    func rename(_ id: UUID, to name: String) {
        guard let i = data.accounts.firstIndex(where: { $0.id == id }) else { return }
        data.accounts[i].name = name
    }

    func deleteAccount(_ id: UUID) {
        var d = data
        d.accounts.removeAll { $0.id == id }
        d.transactions.removeAll { $0.accountID == id }
        d.events.removeAll { $0.accountID == id }
        if d.lifestyle.accountID == id { d.lifestyle.accountID = nil; d.lifestyle.tier = nil }
        if d.lifestyle.debtAccountID == id { d.lifestyle.debtAccountID = nil }
        if d.challenge.accountID == id { d.challenge.accountID = nil }
        d.goals.removeAll { $0.accountID == id }
        d.events.removeAll { $0.toAccountID == id }
        for i in d.accounts.indices where d.accounts[i].terms?.linkedAccountID == id {
            d.accounts[i].terms?.linkedAccountID = nil
        }
        data = d
        syncNotifications()
    }

    // MARK: Transfers (mock)

    enum TransferTarget { case own(UUID), external(name: String, bank: String) }

    func transfer(from: UUID, to target: TransferTarget, amount: Decimal, memo: String) -> Bool {
        guard amount > 0, let src = account(from), src.balance >= amount else { return false }
        var d = data
        switch target {
        case .own(let dst):
            guard let dstAcc = account(dst), dst != from else { return false }
            Ledger.apply(amount: -amount, to: from, date: now, counterparty: dstAcc.name,
                         memo: memo.isEmpty ? "내 계좌 이체" : memo, category: .transfer, tag: SpendingCategory.transfer.rawValue, data: &d)
            Ledger.apply(amount: amount, to: dst, date: now, counterparty: src.name,
                         memo: memo.isEmpty ? "내 계좌 이체" : memo, category: .transfer, tag: SpendingCategory.transfer.rawValue, data: &d)
        case .external(let name, let bank):
            Ledger.apply(amount: -amount, to: from, date: now, counterparty: name,
                         memo: memo.isEmpty ? bank : "\(bank) · \(memo)", category: .transfer, data: &d)
        }
        data = d
        return true
    }

    // MARK: Scheduled events

    func addEvent(_ ev: ScheduledEvent) { data.events.append(ev); syncNotifications() }
    func updateEvent(_ ev: ScheduledEvent) {
        guard let i = data.events.firstIndex(where: { $0.id == ev.id }) else { return }
        data.events[i] = ev; syncNotifications()
    }
    func deleteEvent(_ id: UUID) { data.events.removeAll { $0.id == id }; syncNotifications() }
    func toggleEvent(_ id: UUID) {
        guard let i = data.events.firstIndex(where: { $0.id == id }) else { return }
        data.events[i].isEnabled.toggle()
        // re-enabling a past one-off: push it to the same time tomorrow
        if data.events[i].isEnabled, data.events[i].nextFireDate <= now, data.events[i].repeatRule == .none,
           let t = Calendar.current.date(byAdding: .day, value: 1, to: data.events[i].nextFireDate) {
            data.events[i].nextFireDate = t
        }
        syncNotifications()
    }

    // MARK: Lifestyle (wealth tier)

    var lifestyle: LifestyleState { data.lifestyle }

    /// Pick a wealth tier. Optionally resets the account balance to the tier's
    /// typical amount and back-fills the last `backfillDays` of transactions.
    func applyTier(_ tier: WealthTier?, accountID: UUID?, resetBalance: Bool, backfillDays: Int) {
        if let tier, !data.game.isUnlocked(tier) { return }      // 중산층~재벌 are earned through quests
        var d = data
        if tier != nil, d.game.birthDate == nil {
            d.game.birthDate = Calendar.current.date(byAdding: .year, value: -d.game.startAge, to: d.clock.now())
        }
        // drop the previous tier's auto-created debt account
        if let old = d.lifestyle.debtAccountID {
            d.accounts.removeAll { $0.id == old }
            d.transactions.removeAll { $0.accountID == old }
            d.lifestyle.debtAccountID = nil
        }
        d.lifestyle.tier = tier
        d.lifestyle.accountID = accountID
        d.lifestyle.seed = UInt64.random(in: 1...UInt64.max)
        if let tier, let accountID {
            let simNow = d.clock.now()
            d.lifestyle.tierStart = Calendar.current.date(byAdding: .day, value: -backfillDays, to: simNow) ?? simNow
            if resetBalance { Ledger.setBalance(tier.startingBalance, of: accountID, date: simNow, data: &d) }
            if tier == .broke {
                var debt = Account(name: "\(d.settings.brand.capital) 신용대출", kind: .loan, number: Account.generateNumber(), balance: 0, colorIndex: 5)
                debt.terms = nil // lifestyle engine manages this one
                d.accounts.append(debt)
                Ledger.apply(amount: 4_800_000, to: debt.id, date: d.lifestyle.tierStart, counterparty: d.settings.brand.capital,
                             memo: "기존 대출 잔액 (연 19.9% + 연체)", category: .transfer, data: &d)
                d.lifestyle.debtAccountID = debt.id
            }
            d.lifestyle.lastProcessed = Calendar.current.date(byAdding: .day, value: -backfillDays, to: simNow) ?? simNow
            if backfillDays > 0 {
                // Re-date the balance reset so it precedes the generated history
                if resetBalance, let i = d.transactions.lastIndex(where: { $0.accountID == accountID && $0.category == .adjustment }) {
                    d.transactions[i].date = d.lifestyle.lastProcessed
                }
                LifestyleEngine.process(now: simNow, data: &d)
            }
        } else {
            d.lifestyle.lastProcessed = d.clock.now()
        }
        data = d
    }

    // MARK: 무지출 챌린지

    var challenge: ChallengeState { data.challenge }
    var challengeAccount: Account? { data.challenge.accountID.flatMap(account) }

    func startChallenge() {
        guard data.challenge.accountID == nil else { return }
        var d = data
        let acc = Account(name: "참은 돈 (무지출 챌린지)", kind: .installmentSavings, number: Account.generateNumber(), balance: 0, colorIndex: 3)
        d.accounts.append(acc)
        d.challenge.accountID = acc.id
        d.challenge.startedAt = d.clock.now()
        data = d
    }

    @discardableResult
    func logResist(kind: ResistKind, title: String, amount: Decimal) -> ResistedPurchase? {
        guard amount > 0 else { return nil }
        if data.challenge.accountID == nil { startChallenge() }
        guard let accID = data.challenge.accountID else { return nil }
        var d = data
        let item = ResistedPurchase(date: now, title: title.isEmpty ? kind.placeholder : title, amount: amount, kind: kind)
        d.challenge.items.append(item)
        Ledger.apply(amount: amount, to: accID, date: item.date, counterparty: "\(kind.emoji) \(item.title)", memo: "참은 돈",
                     category: .transfer, tag: SpendingCategory.transfer.rawValue, data: &d)
        data = d
        unlockAchievements(now: now)
        WidgetBridge.publish(store: self)
        return item
    }

    func deleteResist(_ id: UUID) {
        guard let item = data.challenge.items.first(where: { $0.id == id }), let accID = data.challenge.accountID else { return }
        var d = data
        d.challenge.items.removeAll { $0.id == id }
        Ledger.apply(amount: -item.amount, to: accID, date: now, counterparty: "\(item.kind.emoji) \(item.title)", memo: "기록 삭제",
                     category: .transfer, tag: SpendingCategory.transfer.rawValue, data: &d)
        data = d
    }

    // MARK: Savings goals

    var goals: [SavingsGoal] { data.goals }

    func addGoal(title: String, emoji: String, target: Decimal, deadline: Date?, autoAmount: Decimal, fromAccountID: UUID?, dayOfMonth: Int) {
        var d = data
        let acc = Account(name: "\(emoji) \(title)", kind: .deposit, number: Account.generateNumber(), balance: 0, colorIndex: 2)
        d.accounts.append(acc)
        var goal = SavingsGoal(title: title, emoji: emoji, target: target, accountID: acc.id, deadline: deadline, createdAt: now)
        if autoAmount > 0, let from = fromAccountID {
            let cal = Calendar.current
            var comps = cal.dateComponents([.year, .month], from: now)
            comps.day = min(dayOfMonth, 28); comps.hour = 9
            var fire = cal.date(from: comps) ?? now
            if fire <= now { fire = cal.date(byAdding: .month, value: 1, to: fire) ?? fire }
            let ev = ScheduledEvent(accountID: from, counterparty: acc.name, amount: -autoAmount, memo: "목표 자동이체",
                                    nextFireDate: fire, repeatRule: .monthly, toAccountID: acc.id)
            d.events.append(ev)
            goal.autoEventID = ev.id
        }
        d.goals.append(goal)
        data = d
        syncNotifications()
    }

    func addToGoal(_ goalID: UUID, amount: Decimal, from: UUID) {
        guard let g = data.goals.first(where: { $0.id == goalID }), amount > 0 else { return }
        _ = transfer(from: from, to: .own(g.accountID), amount: amount, memo: "목표 저축")
        var d = data; settleGoals(now: now, data: &d); data = d
        unlockAchievements(now: now)
    }

    func deleteGoal(_ id: UUID, refundTo: UUID?) {
        guard let g = data.goals.first(where: { $0.id == id }) else { return }
        var d = data
        if let ev = g.autoEventID { d.events.removeAll { $0.id == ev } }
        if let acc = d.accounts.first(where: { $0.id == g.accountID }) {
            if let to = refundTo, acc.balance > 0 {
                Ledger.apply(amount: acc.balance, to: to, date: now, counterparty: acc.name, memo: "목표 해지 환급",
                             category: .transfer, tag: SpendingCategory.transfer.rawValue, data: &d)
            }
            d.accounts.removeAll { $0.id == acc.id }
            d.transactions.removeAll { $0.accountID == acc.id }
        }
        d.goals.removeAll { $0.id == id }
        data = d
        syncNotifications()
    }

    // MARK: 인생 게임 (quests / tiers)

    var game: GameState { data.game }
    var pendingQuest: Quest? { data.game.pendingQuestID.flatMap(QuestLibrary.quest) }
    var currentTier: WealthTier { data.lifestyle.tier ?? .broke }
    var netWorth: Decimal { QuestEngine.netWorth(data) }
    var age: Int { data.game.age(at: now) }

    struct QuestResult: Identifiable {
        let id = UUID()
        let title: String
        let text: String
        let cash: Decimal
        let delayed: Bool
        let delayMonths: Int
    }

    /// Apply the chosen option of the pending quest.
    @discardableResult
    func answerQuest(_ optionIndex: Int) -> QuestResult? {
        guard let q = pendingQuest, q.options.indices.contains(optionIndex) else { return nil }
        let opt = q.options[optionIndex]
        let e = opt.effect
        var d = data
        let simNow = d.clock.now()
        let cal = Calendar.current
        guard let accID = d.lifestyle.accountID ?? d.accounts.first(where: { $0.kind == .checking })?.id else { return nil }
        let brand = d.settings.brand
        var lines: [String] = []
        if !opt.note.isEmpty { lines.append(opt.note) }
        if let t = e.afterText { lines.append(t) }

        if e.cash != 0 {
            Ledger.apply(amount: e.cash, to: accID, date: simNow, counterparty: brand.resolve(e.employer ?? q.category.title), memo: "[선택] \(String(q.title.prefix(28)))",
                         category: .transfer, tag: (e.cash > 0 ? SpendingCategory.income : categoryTag(q.category)).rawValue, data: &d)
        }
        if let memo = e.cancelFlowMemo { d.game.flows.removeAll { $0.memo.contains(memo) || $0.counterparty.contains(memo) } }
        for f in [e.flow, e.flow2].compactMap({ $0 }) { d.game.flows.append(activeFlow(f, questID: q.id, from: simNow)) }
        if e.salaryMultiplier != 1 { d.game.salaryMultiplier *= e.salaryMultiplier }
        if let emp = e.employer { d.game.employer = emp }
        if let job = e.jobTitle { d.game.jobTitle = job }
        if e.salaryPauseMonths > 0 { d.game.salaryPausedUntil = cal.date(byAdding: .month, value: e.salaryPauseMonths, to: simNow) }
        if e.expenseMultiplier != 1 { d.game.expenseMultiplier *= e.expenseMultiplier }
        if e.medicalMultiplier != 1 { d.game.medicalMultiplier *= e.medicalMultiplier }
        if let m = e.netWorthMultiplier, let acc = d.accounts.first(where: { $0.id == accID }), acc.balance > 0 {
            Ledger.apply(amount: -(acc.balance * Decimal(double: 1 - m)).rounded(), to: accID, date: simNow, counterparty: "재산 분할",
                         memo: "[선택] \(String(q.title.prefix(28)))", category: .transfer, tag: SpendingCategory.other.rawValue, data: &d)
        }
        if let t = e.promoteTo { promote(to: t, data: &d, at: simNow) }
        if e.resist > 0 { data = d; logResist(kind: .shopping, title: String(q.title.prefix(20)), amount: e.resist); d = data }

        var answer = QuestAnswer(questID: q.id, optionIndex: optionIndex, date: simNow, title: q.title, choice: opt.title, outcome: nil, cash: e.cash)
        var delayed = false
        var delayMonths = 0
        var rng = SplitMix64(seed: UInt64.random(in: 1...UInt64.max))
        if let g = e.gamble {
            let success = Double.random(in: 0..<1, using: &rng) < g.chance
            delayMonths = g.delayMonths; delayed = true
            let fire = cal.date(byAdding: .month, value: g.delayMonths, to: simNow) ?? simNow
            var o = PendingOutcome(answerID: answer.id, questID: q.id, fireDate: fire, success: success, cash: success ? g.win : g.lose,
                                   text: success ? g.winText : g.loseText, counterparty: g.counterparty)
            if success {
                if let f = g.winFlow { o.flow = activeFlow(f, questID: q.id, from: fire); o.flowMonths = f.months }
                if let m = e.counterSalary ?? e.successSalary { o.salaryMultiplier = m }
            } else {
                if let f = g.loseFlow { o.flow = activeFlow(f, questID: q.id, from: fire); o.flowMonths = f.months }
                o.salaryPause = max(g.loseSalaryPause, e.loseSalaryPause)
                o.fine = g.loseFine
                o.netWorthMultiplier = e.loseNetWorthMultiplier
                o.demoteTo = e.loseDemoteTo
            }
            if o.text.isEmpty { o.text = success ? "별다른 일 없이 지나갔습니다." : "안 좋은 결과가 나왔습니다." }
            d.game.outcomes.append(o)
            lines.append("결과는 \(g.delayMonths)개월 뒤에 알 수 있습니다.")
        }
        if let v = e.venture {
            let success = Double.random(in: 0..<1, using: &rng) < v.chance
            delayMonths = max(delayMonths, v.delayMonths); delayed = true
            let fire = cal.date(byAdding: .month, value: v.delayMonths, to: simNow) ?? simNow
            var o = PendingOutcome(answerID: answer.id, questID: q.id, fireDate: fire, success: success, cash: 0,
                                   text: success ? v.winText : v.loseText, counterparty: v.winFlow.template.counterparty)
            if success { o.flow = activeFlow(v.winFlow, questID: q.id, from: fire); o.flowMonths = v.winFlow.months }
            else if let lf = v.loseFlow { o.flow = activeFlow(lf, questID: q.id, from: fire); o.flowMonths = lf.months }
            d.game.outcomes.append(o)
            lines.append("사업 결과는 \(v.delayMonths)개월 뒤에 드러납니다.")
        }
        if !delayed { answer.outcome = lines.joined(separator: " ") }
        d.game.answers.append(answer)
        d.game.pendingQuestID = nil; d.game.pendingSince = nil; d.game.pendingExpires = nil
        QuestEngine.scheduleNext(from: simNow, tier: d.lifestyle.tier ?? .broke, age: d.game.age(at: simNow), netWorth: QuestEngine.netWorth(d), rng: &rng, data: &d)
        data = d
        syncNotifications()
        unlockAchievements(now: now)
        WidgetBridge.publish(store: self)
        let result = QuestResult(title: opt.title, text: lines.joined(separator: "\n"), cash: e.cash, delayed: delayed, delayMonths: delayMonths)
        questResult = result
        return result
    }

    /// Answer from a notification action: the quest must be the pending one (catchUp ran first).
    func handleQuestNotification(questID: String, optionIndex: Int?) {
        catchUp()
        guard data.game.pendingQuestID == questID else { return }
        if let i = optionIndex { answerQuest(i) }
    }

    /// Test/UI-test hook: make a specific quest the pending one right now.
    func forcePendingQuest(_ id: String) {
        guard QuestLibrary.quest(id) != nil else { return }
        var d = data
        if d.lifestyle.tier == nil, let acc = d.accounts.first(where: { $0.kind == .checking }) {
            d.lifestyle.tier = .working; d.lifestyle.accountID = acc.id; d.lifestyle.lastProcessed = d.clock.now()
        }
        if d.game.birthDate == nil { d.game.birthDate = Calendar.current.date(byAdding: .year, value: -d.game.startAge, to: d.clock.now()) }
        d.game.pendingQuestID = id; d.game.pendingSince = d.clock.now()
        d.game.pendingExpires = Calendar.current.date(byAdding: .day, value: GameState.questWindowDays, to: d.clock.now())
        data = d
    }

    func dismissQuest() {
        var d = data
        d.game.pendingQuestID = nil; d.game.pendingSince = nil; d.game.pendingExpires = nil
        d.game.nextQuestDate = Calendar.current.date(byAdding: .day, value: 2, to: d.clock.now())
        data = d
        syncNotifications()
    }

    private func activeFlow(_ f: QuestFlow, questID: String, from: Date) -> ActiveFlow {
        let t = f.template
        var day = 25
        if case .monthly(let d) = t.frequency { day = d }
        return ActiveFlow(questID: questID, counterparty: t.counterparty, memo: t.memo, amount: t.minAmount, day: day, category: t.category, hour: t.hour,
                          until: f.months.flatMap { Calendar.current.date(byAdding: .month, value: $0, to: from) })
    }

    private func categoryTag(_ c: QuestCategory) -> SpendingCategory {
        switch c {
        case .insurance: return .insurance
        case .housing: return .housing
        case .invest, .finance: return .finance
        case .business: return .other
        case .family: return .family
        case .education: return .education
        case .health: return .medical
        case .luxury: return .shopping
        case .tax: return .tax
        default: return .other
        }
    }

    /// Tier change by quest: keeps balances, unlocks the tier, restarts raise/inflation anchors.
    private func promote(to tier: WealthTier, data d: inout BankData, at date: Date) {
        if !d.game.unlockedTiers.contains(tier) { d.game.unlockedTiers.append(tier) }
        d.lifestyle.tier = tier
        d.lifestyle.tierStart = date
        d.lifestyle.seed = UInt64.random(in: 1...UInt64.max)
        if tier != .broke, let debt = d.lifestyle.debtAccountID {
            d.accounts.removeAll { $0.id == debt }; d.transactions.removeAll { $0.accountID == debt }; d.lifestyle.debtAccountID = nil
        }
    }

    func restartLife() {
        resetAll()
    }

    // MARK: Scenarios

    func applyScenario(_ scenario: Scenario) {
        if checkingAccounts.isEmpty { addChecking(name: "입출금통장", balance: 0, colorIndex: 0) }
        guard let accID = checkingAccounts.first?.id else { return }
        applyTier(scenario.tier, accountID: accID, resetBalance: true, backfillDays: 30)
        var d = data
        if let bal = scenario.startingBalance { Ledger.setBalance(bal, of: accID, date: d.clock.now(), data: &d) }
        let simNow = d.clock.now()
        for e in scenario.events {
            let fire = simNow.addingTimeInterval(TimeInterval(e.minutesFromNow * 60))
            d.events.append(ScheduledEvent(accountID: accID, counterparty: d.settings.brand.resolve(e.counterparty), amount: e.amount,
                                           memo: e.memo, nextFireDate: fire, repeatRule: e.repeatRule))
        }
        data = d
        syncNotifications()
    }

    // MARK: Mock trading

    func updatePortfolio(_ change: (inout Portfolio) -> Void) {
        var p = data.portfolio; change(&p); data.portfolio = p
    }

    func addHolding(_ h: Holding) { updatePortfolio { $0.holdings.append(h) } }

    func removeHolding(_ id: UUID) {
        updatePortfolio { $0.holdings.removeAll { $0.id == id }; $0.trades.removeAll { $0.holdingID == id } }
    }

    func setPrice(_ price: Decimal, of id: UUID) {
        updatePortfolio { p in
            if let i = p.holdings.firstIndex(where: { $0.id == id }) { p.holdings[i].price = price }
        }
    }

    /// Dry-run order. Buys draw from mock cash in the holding's currency; sells add to it.
    func trade(holdingID: UUID, buy: Bool, shares: Decimal, price: Decimal) -> Bool {
        guard shares > 0, price > 0 else { return false }
        var p = data.portfolio
        guard let i = p.holdings.firstIndex(where: { $0.id == holdingID }) else { return false }
        let total = (shares * price).rounded(scale: 2)
        var h = p.holdings[i]
        if buy {
            if h.market == .domestic { guard p.cashKRW >= total else { return false }; p.cashKRW -= total }
            else { guard p.cashUSD >= total else { return false }; p.cashUSD -= total }
            h.avgCost = ((h.cost + total) / (h.shares + shares)).rounded(scale: 2)
            h.shares += shares
        } else {
            guard h.shares >= shares else { return false }
            if h.market == .domestic { p.cashKRW += total } else { p.cashUSD += total }
            h.shares -= shares
        }
        h.price = price
        p.holdings[i] = h
        p.trades.append(TradeRecord(holdingID: holdingID, date: now, isBuy: buy, shares: shares, price: price))
        data.portfolio = p
        return true
    }

    // MARK: Products

    func openProduct(kind: AccountKind, name: String, annualRate: Double, termMonths: Int,
                     amount: Decimal, linkedAccountID: UUID?, startDate: Date) {
        var d = data
        let number = Account.generateNumber()
        var terms = ProductTerms(annualRate: annualRate, termMonths: termMonths, startDate: startDate,
                                 principal: 0, monthlyAmount: 0, linkedAccountID: linkedAccountID,
                                 lastAccrualDate: startDate, lastInstallmentDate: startDate)
        var acc = Account(name: name, kind: kind, number: number, balance: 0, terms: nil, colorIndex: Int.random(in: 0..<6))
        switch kind {
        case .deposit:
            terms.principal = amount
            d.accounts.append(acc)
            if let linked = linkedAccountID { // move the lump sum out of checking
                Ledger.apply(amount: -amount, to: linked, date: startDate, counterparty: name, memo: "예금 가입", category: .transfer, data: &d)
            }
            Ledger.apply(amount: amount, to: acc.id, date: startDate, counterparty: "가입", memo: "예치금", category: .transfer, data: &d)
        case .installmentSavings:
            terms.monthlyAmount = amount
            d.accounts.append(acc)
            // first installment on the start date
            if let linked = linkedAccountID {
                Ledger.apply(amount: -amount, to: linked, date: startDate, counterparty: name, memo: "적금 자동이체", category: .autoTransfer, data: &d)
            }
            Ledger.apply(amount: amount, to: acc.id, date: startDate, counterparty: "1회차", memo: "월 납입", category: .autoTransfer, data: &d)
        case .loan:
            terms.principal = amount
            terms.monthlyAmount = InterestEngine.annuityPayment(principal: amount, annualRate: annualRate, months: termMonths)
            acc.balance = 0
            d.accounts.append(acc)
            Ledger.apply(amount: amount, to: acc.id, date: startDate, counterparty: "대출 실행", memo: "원금", category: .transfer, data: &d)
            if let linked = linkedAccountID {
                Ledger.apply(amount: amount, to: linked, date: startDate, counterparty: name, memo: "대출금 입금", category: .transfer, data: &d)
            }
        case .checking:
            return
        }
        if let i = d.accounts.firstIndex(where: { $0.id == acc.id }) { d.accounts[i].terms = terms }
        data = d
        catchUp() // if startDate is in the past, settle up immediately
    }
}

import SwiftUI

/// 저축 목표: a dedicated account per goal, optional monthly auto-transfer that feeds it.
struct GoalsView: View {
    @EnvironmentObject var store: BankStore
    @State private var adding = false
    @State private var funding: SavingsGoal?

    var body: some View {
        ScrollView {
            VStack(spacing: 14) {
                if store.goals.isEmpty {
                    ContentUnavailableView("저축 목표가 없어요", systemImage: "target", description: Text("제주 여행, 맥북, 비상금처럼 목표를 만들고 매달 자동이체로 채워 보세요."))
                        .padding(.top, 40)
                }
                ForEach(store.goals) { g in GoalCard(goal: g) { funding = g } }
                Button { adding = true } label: {
                    Label("목표 만들기", systemImage: "plus").font(.app(.body, .bold)).frame(maxWidth: .infinity).padding(.vertical, 12)
                }.buttonStyle(.borderedProminent)
            }
            .padding()
        }
        .background(store.theme.background)
        .navigationTitle("저축 목표")
        .navigationBarTitleDisplayMode(.inline)
        .sheet(isPresented: $adding) { GoalEditor() }
        .sheet(item: $funding) { FundGoalSheet(goal: $0) }
    }
}

struct GoalCard: View {
    @EnvironmentObject var store: BankStore
    let goal: SavingsGoal
    let fund: () -> Void
    @State private var confirmDelete = false

    var body: some View {
        let balance = store.account(goal.accountID)?.balance ?? 0
        let p = goal.progress(balance: balance)
        Card {
            VStack(alignment: .leading, spacing: 10) {
                HStack {
                    Text(goal.emoji).font(.app(.title2))
                    VStack(alignment: .leading, spacing: 2) {
                        Text(goal.title).font(.app(.headline, .semibold))
                        if let d = goal.deadline { Text("목표일 \(d.formatted(date: .abbreviated, time: .omitted))").font(.app(.caption)).foregroundStyle(.secondary) }
                    }
                    Spacer()
                    if goal.reachedAt != nil { Label("달성", systemImage: "checkmark.seal.fill").font(.app(.caption, .semibold)).foregroundStyle(.green) }
                    Menu {
                        Button(role: .destructive) { confirmDelete = true } label: { Label("목표 삭제 (잔액 환급)", systemImage: "trash") }
                    } label: { Image(systemName: "ellipsis.circle").foregroundStyle(.secondary) }
                }
                HStack(alignment: .firstTextBaseline) {
                    MoneyText(amount: balance, font: .app(.title3, .semibold))
                    Text("/ \(Money.format(goal.target, code: store.currency))").font(.app(.caption)).foregroundStyle(.secondary)
                    Spacer()
                    Text("\(Int(p * 100))%").font(.app(.subheadline, .bold)).foregroundStyle(store.theme.accent)
                }
                GeometryReader { geo in
                    Capsule().fill(Color.primary.opacity(0.07))
                        .overlay(alignment: .leading) { Capsule().fill(store.theme.accent).frame(width: geo.size.width * p) }
                }.frame(height: 8)
                HStack {
                    if let evID = goal.autoEventID, let ev = store.data.events.first(where: { $0.id == evID }) {
                        Label("매월 \(Calendar.current.component(.day, from: ev.nextFireDate))일 \(Money.format(abs(ev.amount), code: store.currency)) 자동이체", systemImage: "repeat")
                            .font(.app(.caption)).foregroundStyle(.secondary)
                    }
                    Spacer()
                    Button("채우기") { fund() }.font(.app(.subheadline, .semibold))
                }
            }
        }
        .confirmationDialog("목표를 삭제할까요? 모은 돈은 입출금통장으로 돌아갑니다.", isPresented: $confirmDelete, titleVisibility: .visible) {
            Button("삭제", role: .destructive) { store.deleteGoal(goal.id, refundTo: store.checkingAccounts.first?.id) }
        }
    }
}

struct GoalEditor: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    @State private var title = ""
    @State private var emoji = "✈️"
    @State private var target: Decimal = 3_000_000
    @State private var hasDeadline = false
    @State private var deadline = Calendar.current.date(byAdding: .month, value: 6, to: Date()) ?? Date()
    @State private var auto = true
    @State private var autoAmount: Decimal = 300_000
    @State private var fromID: UUID?
    @State private var day = 25

    private let emojis = ["✈️", "💻", "🏠", "🚗", "💍", "🎓", "🏖️", "🛡️", "🎁", "📱", "🐶", "💪"]
    private let presets: [(String, String, Decimal)] = [("제주 여행", "✈️", 1_500_000), ("맥북", "💻", 2_800_000), ("비상금", "🛡️", 5_000_000), ("전세 보증금", "🏠", 50_000_000)]

    var body: some View {
        NavigationStack {
            Form {
                Section("빠른 선택") {
                    ScrollView(.horizontal, showsIndicators: false) {
                        HStack {
                            ForEach(presets, id: \.0) { p in
                                Button { title = p.0; emoji = p.1; target = p.2 } label: {
                                    Text("\(p.1) \(p.0)").font(.app(.caption, .semibold)).padding(10)
                                        .background(store.theme.accent.opacity(0.1), in: RoundedRectangle(cornerRadius: 10))
                                }.buttonStyle(.plain)
                            }
                        }
                    }.listRowInsets(EdgeInsets(top: 8, leading: 16, bottom: 8, trailing: 16))
                }
                Section("목표") {
                    HStack(spacing: 8) {
                        ForEach(emojis, id: \.self) { e in
                            Text(e).font(.app(.title3)).padding(4)
                                .background(emoji == e ? store.theme.accent.opacity(0.2) : .clear, in: RoundedRectangle(cornerRadius: 8))
                                .onTapGesture { emoji = e }
                        }
                    }
                    TextField("목표 이름 (예: 제주 여행)", text: $title)
                    AmountField(title: "목표 금액", amount: $target)
                    Toggle("목표일", isOn: $hasDeadline)
                    if hasDeadline { DatePicker("날짜", selection: $deadline, displayedComponents: .date) }
                }
                Section {
                    Toggle("매월 자동이체", isOn: $auto)
                    if auto {
                        AmountField(title: "금액", amount: $autoAmount)
                        Picker("출금 계좌", selection: $fromID) {
                            ForEach(store.checkingAccounts) { a in Text(a.name).tag(Optional(a.id)) }
                        }
                        Stepper("매월 \(day)일", value: $day, in: 1...28)
                    }
                } footer: {
                    if auto, autoAmount > 0, target > 0 {
                        Text("이 속도면 약 \(Int(((target / autoAmount).doubleValue).rounded(.up)))개월 뒤 달성합니다.")
                    }
                }
            }
            .navigationTitle("목표 만들기")
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .cancellationAction) { Button("취소") { dismiss() } }
                ToolbarItem(placement: .confirmationAction) {
                    Button("만들기") {
                        store.addGoal(title: title, emoji: emoji, target: target, deadline: hasDeadline ? deadline : nil,
                                      autoAmount: auto ? autoAmount : 0, fromAccountID: fromID, dayOfMonth: day)
                        dismiss()
                    }.disabled(title.isEmpty || target <= 0)
                }
            }
            .onAppear { fromID = store.checkingAccounts.first?.id }
        }
    }
}

struct FundGoalSheet: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    let goal: SavingsGoal
    @State private var amount: Decimal = 100_000
    @State private var fromID: UUID?

    var body: some View {
        NavigationStack {
            Form {
                AmountField(title: "금액", amount: $amount)
                Picker("출금 계좌", selection: $fromID) { ForEach(store.checkingAccounts) { a in Text(a.name).tag(Optional(a.id)) } }
            }
            .navigationTitle("\(goal.emoji) \(goal.title) 채우기")
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .cancellationAction) { Button("취소") { dismiss() } }
                ToolbarItem(placement: .confirmationAction) {
                    Button("이체") { if let f = fromID { store.addToGoal(goal.id, amount: amount, from: f) }; dismiss() }
                        .disabled(amount <= 0 || fromID == nil)
                }
            }
            .onAppear { fromID = store.checkingAccounts.first?.id }
        }
    }
}

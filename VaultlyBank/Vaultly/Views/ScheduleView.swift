import SwiftUI

/// "엄마가 오후 4시에 ₩5,000,000 용돈을 보냄" 같은 예약 입금/출금 알림.
struct ScheduleView: View {
    @EnvironmentObject var store: BankStore
    @State private var editing: ScheduledEvent?
    @State private var creating = false

    var body: some View {
        NavigationStack {
            List {
                if !store.notificationsAuthorized {
                    Section {
                        Button { store.requestNotifications() } label: {
                            Label("알림 권한 허용하기", systemImage: "bell.badge.fill")
                        }
                    } footer: { Text("권한을 허용하면 예약 시각에 잠금화면 알림이 표시됩니다. 권한이 없어도 잔액은 예약대로 반영됩니다.") }
                }
                let events = store.data.events.sorted { $0.nextFireDate < $1.nextFireDate }
                if events.isEmpty {
                    ContentUnavailableView("예약된 알림이 없습니다", systemImage: "bell.slash",
                                           description: Text("+ 버튼으로 입금 알림을 예약해 보세요."))
                }
                ForEach(events) { ev in
                    Button { editing = ev } label: { EventRow(event: ev) }
                        .buttonStyle(.plain)
                        .swipeActions {
                            Button(role: .destructive) { store.deleteEvent(ev.id) } label: { Label("삭제", systemImage: "trash") }
                            Button { store.toggleEvent(ev.id) } label: {
                                Label(ev.isEnabled ? "끄기" : "켜기", systemImage: ev.isEnabled ? "pause" : "play")
                            }.tint(.orange)
                        }
                }
            }
            .navigationTitle("예약 알림")
            .toolbar {
                ToolbarItem(placement: .topBarLeading) { SimBadge() }
                ToolbarItem(placement: .topBarTrailing) { Button { creating = true } label: { Image(systemName: "plus") } }
            }
            .sheet(isPresented: $creating) { EventEditor(event: nil) }
            .sheet(item: $editing) { EventEditor(event: $0) }
        }
    }
}

struct EventRow: View {
    @EnvironmentObject var store: BankStore
    let event: ScheduledEvent
    var body: some View {
        HStack(spacing: 12) {
            Image(systemName: event.isIncoming ? "arrow.down.left.circle.fill" : "arrow.up.right.circle.fill")
                .font(.title2).foregroundStyle(event.isEnabled ? (event.isIncoming ? .green : .orange) : .gray)
            VStack(alignment: .leading, spacing: 2) {
                Text("\(event.counterparty) · \(event.memo)").font(.subheadline.weight(.semibold))
                Text(schedule).font(.caption).foregroundStyle(.secondary)
                if let a = store.account(event.accountID) { Text(a.name).font(.caption2).foregroundStyle(.tertiary) }
            }
            Spacer()
            MoneyText(amount: event.amount, signed: true, font: .subheadline.weight(.bold), hideable: false)
                .foregroundStyle(event.isEnabled ? .primary : .secondary)
        }
        .opacity(event.isEnabled ? 1 : 0.6)
    }
    private var schedule: String {
        let when = event.nextFireDate.formatted(date: .abbreviated, time: .shortened)
        if !event.isEnabled { return "꺼짐 · 마지막 \(event.lastFiredAt?.formatted(date: .abbreviated, time: .shortened) ?? "-")" }
        return event.repeatRule == .none ? when : "\(event.repeatRule.title) · 다음 \(when)"
    }
}

struct EventEditor: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    let event: ScheduledEvent?

    @State private var counterparty = ""
    @State private var memo = ""
    @State private var incoming = true
    @State private var amount: Decimal = 0
    @State private var accountID: UUID?
    @State private var date = Calendar.current.date(bySettingHour: 16, minute: 0, second: 0, of: Date()) ?? Date()
    @State private var rule: RepeatRule = .none

    private let presets: [(String, String, Decimal)] = [
        ("엄마", "용돈", 5_000_000), ("아빠", "생활비", 1_000_000), ("회사", "급여", 4_500_000), ("고객", "정산", 12_000_000),
    ]

    var body: some View {
        NavigationStack {
            Form {
                Section("빠른 선택") {
                    ScrollView(.horizontal, showsIndicators: false) {
                        HStack {
                            ForEach(presets, id: \.0) { p in
                                Button {
                                    counterparty = p.0; memo = p.1; amount = p.2; incoming = true
                                } label: {
                                    VStack(alignment: .leading) {
                                        Text("\(p.0) · \(p.1)").font(.caption.weight(.semibold))
                                        Text(Money.format(p.2, code: store.currency)).font(.caption2)
                                    }
                                    .padding(10).background(store.theme.accent.opacity(0.1), in: RoundedRectangle(cornerRadius: 10))
                                }.buttonStyle(.plain)
                            }
                        }
                    }.listRowInsets(EdgeInsets(top: 8, leading: 16, bottom: 8, trailing: 16))
                }
                Section("내용") {
                    Picker("종류", selection: $incoming) { Text("입금").tag(true); Text("출금").tag(false) }.pickerStyle(.segmented)
                    TextField(incoming ? "보내는 사람 (예: 엄마)" : "받는 곳 (예: 카드 대금)", text: $counterparty)
                    TextField("메모 (예: 용돈)", text: $memo)
                    AmountField(title: "금액", amount: $amount)
                    Picker("계좌", selection: $accountID) {
                        ForEach(store.checkingAccounts) { a in Text(a.name).tag(Optional(a.id)) }
                    }
                }
                Section("시각") {
                    DatePicker("날짜와 시간", selection: $date)
                    Picker("반복", selection: $rule) { ForEach(RepeatRule.allCases) { Text($0.title).tag($0) } }
                }
                if event != nil {
                    Section {
                        Button("삭제", role: .destructive) { store.deleteEvent(event!.id); dismiss() }
                    }
                }
            }
            .navigationTitle(event == nil ? "알림 예약" : "알림 수정")
            .toolbar {
                ToolbarItem(placement: .cancellationAction) { Button("취소") { dismiss() } }
                ToolbarItem(placement: .confirmationAction) { Button("저장") { save() }.disabled(!valid) }
            }
            .onAppear {
                if let e = event {
                    counterparty = e.counterparty; memo = e.memo; incoming = e.isIncoming; amount = abs(e.amount)
                    accountID = e.accountID; date = e.nextFireDate; rule = e.repeatRule
                } else {
                    accountID = store.checkingAccounts.first?.id
                    date = Calendar.current.date(bySettingHour: 16, minute: 0, second: 0, of: store.now) ?? store.now
                    if date <= store.now, let t = Calendar.current.date(byAdding: .day, value: 1, to: date) { date = t }
                }
            }
        }
    }

    private var valid: Bool { amount > 0 && !counterparty.isEmpty && accountID != nil }

    private func save() {
        guard let accountID else { return }
        let signed = incoming ? amount : -amount
        var ev = event ?? ScheduledEvent(accountID: accountID, counterparty: counterparty, amount: signed, memo: memo,
                                         nextFireDate: date, repeatRule: rule)
        ev.accountID = accountID; ev.counterparty = counterparty; ev.amount = signed; ev.memo = memo
        ev.nextFireDate = date; ev.repeatRule = rule; ev.isEnabled = true
        if event == nil { store.addEvent(ev) } else { store.updateEvent(ev) }
        store.catchUp() // a time already in the past applies immediately
        dismiss()
    }
}

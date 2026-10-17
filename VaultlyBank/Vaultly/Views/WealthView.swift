import SwiftUI

/// Choose a wealth tier; the app then keeps generating matching cash flows.
struct WealthView: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    @State private var tier: WealthTier?
    @State private var accountID: UUID?
    @State private var resetBalance = true
    @State private var backfill = true

    var body: some View {
        NavigationStack {
            Form {
                Section {
                    ForEach(WealthTier.allCases) { t in
                        let locked = !store.game.isUnlocked(t)
                        Button { if !locked { tier = t } } label: {
                            HStack(alignment: .top, spacing: 12) {
                                Text(t.emoji).font(.app(.title)).opacity(locked ? 0.4 : 1)
                                VStack(alignment: .leading, spacing: 4) {
                                    HStack {
                                        Text(t.title).font(.app(.headline, .semibold))
                                        Text("잔액 \(Money.format(t.startingBalance, code: store.currency))").font(.app(.caption)).foregroundStyle(.secondary)
                                    }
                                    Text(locked ? "잠김 · 순자산 \(Money.format(GameState.threshold(for: t), code: store.currency)) 이상에서 승급 퀘스트로 해제" : t.summary).font(.app(.caption)).foregroundStyle(.secondary)
                                }
                                Spacer()
                                if locked { Image(systemName: "lock.fill").foregroundStyle(.tertiary) }
                                else if tier == t { Image(systemName: "checkmark.circle.fill").foregroundStyle(store.theme.accent) }
                            }
                            .contentShape(Rectangle())
                        }.buttonStyle(.plain)
                    }
                    Button(role: .destructive) { tier = nil } label: {
                        HStack { Text("자동 흐름 끄기"); Spacer(); if tier == nil { Image(systemName: "checkmark") } }
                    }
                } header: { Text("재력 정도") } footer: {
                    Text("거지·서민에서 시작합니다. 중산층·부자·재벌은 순자산을 모으고 승급 퀘스트(이직·청약·스톡옵션·상장)를 통과해야 열립니다. 선택한 재력에 맞는 급여·지출·보험료·공과금·의료비·사고·경조사·세금·보너스가 자동으로 생기고, 며칠마다 인생을 바꾸는 선택지가 옵니다.")
                }
                if tier != nil {
                    Section("적용 대상") {
                        Picker("계좌", selection: $accountID) {
                            ForEach(store.checkingAccounts) { a in Text(a.name).tag(Optional(a.id)) }
                        }
                        Toggle("잔액도 재력에 맞게 설정", isOn: $resetBalance)
                        Toggle("지난 30일 내역 바로 생성", isOn: $backfill)
                    }
                    if let t = tier {
                        Section("현실 반영") {
                            HStack { Text("급여 인상"); Spacer(); Text(String(format: "매년 +%.1f%% · %d년마다 승진 ×%.2f", t.annualRaise * 100, t.promotionEveryYears, t.promotionBump)).font(.app(.caption)).foregroundStyle(.secondary) }
                            HStack { Text("물가 상승"); Spacer(); Text("지출 매년 +2.5%").font(.app(.caption)).foregroundStyle(.secondary) }
                            HStack { Text("사건·사고 변형"); Spacer(); Text("\(EventLibrary.variationCount(for: t))가지").font(.app(.caption)).foregroundStyle(.secondary) }
                            if t == .broke {
                                Text("잔액이 바닥나면 캐피탈에서 자동으로 추가 대출이 실행되고, 매월 이자·연체 수수료가 원금에 붙어 빚이 계속 불어납니다.")
                                    .font(.app(.caption)).foregroundStyle(.red)
                            }
                        }
                        Section("기본 흐름") {
                            ForEach(Array(t.flows.enumerated()), id: \.offset) { _, f in
                                HStack {
                                    VStack(alignment: .leading, spacing: 2) {
                                        Text("\(f.counterparty) · \(f.memo)").font(.app(.subheadline))
                                        Text(frequencyText(f.frequency)).font(.app(.caption)).foregroundStyle(.secondary)
                                    }
                                    Spacer()
                                    Text(amountText(f)).font(.app(.subheadline, .semibold)).monospacedDigit()
                                        .foregroundStyle(f.minAmount >= 0 ? Color.blue : Color.primary)
                                }
                            }
                        }
                    }
                }
            }
            .navigationTitle("재력 설정")
            .toolbar {
                ToolbarItem(placement: .cancellationAction) { Button("취소") { dismiss() } }
                ToolbarItem(placement: .confirmationAction) {
                    Button("적용") {
                        store.applyTier(tier, accountID: accountID, resetBalance: resetBalance, backfillDays: backfill ? 30 : 0)
                        dismiss()
                    }.disabled(tier != nil && accountID == nil)
                }
            }
            .onAppear {
                tier = store.lifestyle.tier
                accountID = store.lifestyle.accountID ?? store.checkingAccounts.first?.id
            }
        }
    }

    private func frequencyText(_ f: FlowFrequency) -> String {
        switch f {
        case .monthly(let d): return "매월 \(d)일"
        case .weekly(let wd): return "매주 \(["일", "월", "화", "수", "목", "금", "토"][max(0, min(6, wd - 1))])요일"
        case .daily(let c): return "거의 매일 (\(Int(c * 100))%)"
        }
    }
    private func amountText(_ f: FlowTemplate) -> String {
        let code = store.currency
        if f.minAmount == f.maxAmount { return Money.format(f.minAmount, code: code, signed: true) }
        return "\(Money.format(f.minAmount, code: code, signed: true)) ~ \(Money.format(f.maxAmount, code: code, signed: true))"
    }
}

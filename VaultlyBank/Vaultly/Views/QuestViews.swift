import SwiftUI

/// Pending life decision on Home. Each option changes the asset flow.
struct QuestCard: View {
    @EnvironmentObject var store: BankStore
    let quest: Quest
    @State private var result: BankStore.QuestResult?
    @State private var confirmIndex: Int?

    var body: some View {
        Card {
            VStack(alignment: .leading, spacing: 12) {
                HStack(spacing: 8) {
                    Image(systemName: quest.category.symbol).font(.app(.caption, .semibold)).foregroundStyle(store.theme.accent)
                    Text("새로운 선택 · \(quest.category.title)").font(.app(.caption, .semibold)).foregroundStyle(store.theme.accent)
                    Spacer()
                    if let exp = store.game.pendingExpires {
                        let d = max(0, Calendar.current.dateComponents([.day], from: store.now, to: exp).day ?? 0)
                        Text("D-\(d)").font(.app(.caption2, .semibold)).foregroundStyle(.secondary)
                    }
                    Button { store.dismissQuest() } label: { Image(systemName: "xmark").font(.app(.caption)).foregroundStyle(.tertiary) }
                }
                Text(quest.title).font(.app(.headline, .semibold)).fixedSize(horizontal: false, vertical: true)
                if !quest.detail.isEmpty {
                    Text(quest.detail).font(.app(.subheadline)).foregroundStyle(.secondary).fixedSize(horizontal: false, vertical: true)
                }
                ForEach(Array(quest.options.enumerated()), id: \.offset) { i, opt in
                    Button { confirmIndex = i } label: { optionLabel(opt, primary: i == 0) }
                        .buttonStyle(.plain)
                        .accessibilityIdentifier("quest-option-\(i)")
                }
            }
        }
        .confirmationDialog(confirmIndex.map { quest.options[$0].title } ?? "", isPresented: Binding(get: { confirmIndex != nil }, set: { if !$0 { confirmIndex = nil } }), titleVisibility: .visible) {
            Button("이렇게 한다") { if let i = confirmIndex { result = store.answerQuest(i) }; confirmIndex = nil }
            Button("취소", role: .cancel) { confirmIndex = nil }
        } message: { Text("한 번 정하면 되돌릴 수 없습니다.") }
        .sheet(item: Binding(get: { result.map { IdentifiedResult(r: $0) } }, set: { _ in result = nil })) { r in
            OutcomeSheet(result: r.r).presentationDetents([.medium])
        }
    }
    private struct IdentifiedResult: Identifiable { let id = UUID(); let r: BankStore.QuestResult }

    private func optionLabel(_ opt: QuestOption, primary: Bool) -> some View {
        let fg: Color = primary ? store.theme.onAccent : .primary
        let sub: Color = primary ? store.theme.onAccent.opacity(0.8) : .secondary
        let bg: Color = primary ? store.theme.accent : Color.primary.opacity(0.06)
        return HStack {
            VStack(alignment: .leading, spacing: 2) {
                Text(opt.title).font(.app(.body, .semibold)).foregroundStyle(fg)
                if !opt.note.isEmpty { Text(opt.note).font(.app(.caption)).foregroundStyle(sub) }
            }
            Spacer()
            Image(systemName: "chevron.right").font(.app(.caption)).foregroundStyle(sub)
        }
        .padding(12)
        .background(bg, in: RoundedRectangle(cornerRadius: 12, style: .continuous))
    }
}

struct OutcomeSheet: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    let result: BankStore.QuestResult
    var body: some View {
        VStack(spacing: 16) {
            Image(systemName: result.delayed ? "hourglass" : "checkmark.circle.fill").font(.app(size: 44)).foregroundStyle(store.theme.accent).padding(.top, 24)
            Text(result.title).font(.app(.title3, .bold))
            if result.cash != 0 {
                MoneyText(amount: result.cash, signed: true, font: .app(.title2, .semibold), hideable: false)
                    .foregroundStyle(result.cash > 0 ? store.theme.accent : store.theme.accentSecondary)
            }
            if !result.text.isEmpty { Text(result.text).font(.app(.subheadline)).foregroundStyle(.secondary).multilineTextAlignment(.center).padding(.horizontal) }
            if result.delayed { Text("\(result.delayMonths)개월 뒤 결과가 알림으로 옵니다.").font(.app(.caption)).foregroundStyle(.tertiary) }
            Spacer()
            Button { dismiss() } label: { Text("확인").frame(maxWidth: .infinity).padding(.vertical, 6) }.buttonStyle(.borderedProminent).padding()
        }
    }
}

/// 내 인생: age, tier ladder with locks, job, active flows, pending outcomes, decision timeline.
struct LifeView: View {
    @EnvironmentObject var store: BankStore
    @State private var showEnding = false

    var body: some View {
        ScrollView {
            VStack(spacing: 16) {
                header
                if let q = store.pendingQuest { QuestCard(quest: q) }
                ladder
                flows
                timeline
            }
            .padding()
        }
        .background(store.theme.background)
        .navigationTitle("내 인생")
        .navigationBarTitleDisplayMode(.inline)
        .toolbar {
            ToolbarItem(placement: .topBarTrailing) {
                Menu {
                    ForEach(TimeSpeed.allCases) { s in Button(s.title) { store.setSpeed(s) } }
                } label: { Label("속도", systemImage: "forward.fill") }
            }
        }
        .sheet(isPresented: $showEnding) { EndingView() }
        .onAppear { showEnding = store.game.isEnded }
        .onChange(of: store.game.isEnded) { _, ended in showEnding = ended }
    }

    private var header: some View {
        let tier = store.currentTier
        return Card {
            VStack(alignment: .leading, spacing: 8) {
                HStack {
                    Text("\(tier.emoji) \(tier.title)").font(.app(.title3, .bold))
                    Spacer()
                    Text("만 \(store.age)세").font(.app(.subheadline, .semibold)).foregroundStyle(.secondary)
                }
                Text("순자산").font(.app(.caption)).foregroundStyle(.secondary)
                MoneyText(amount: store.netWorth, font: .app(size: 28, .semibold))
                if let emp = store.game.employer {
                    Label("\(emp)\(store.game.jobTitle.map { " · \($0)" } ?? "")", systemImage: "briefcase").font(.app(.caption)).foregroundStyle(.secondary).lineLimit(1)
                }
                HStack(spacing: 12) {
                    stat("연봉 배수", String(format: "×%.2f", store.game.salaryMultiplier))
                    stat("지출 배수", String(format: "×%.2f", store.game.expenseMultiplier))
                    stat("병원비 배수", String(format: "×%.2f", store.game.medicalMultiplier))
                    stat("선택", "\(store.game.answers.count)")
                }
                if let until = store.game.salaryPausedUntil, until > store.now {
                    Label("급여 중단 · \(until.formatted(date: .abbreviated, time: .omitted))까지", systemImage: "pause.circle").font(.app(.caption)).foregroundStyle(store.theme.accentSecondary)
                }
                HStack {
                    Image(systemName: "forward.fill").foregroundStyle(store.theme.accent)
                    Text(TimeSpeed.allCases.first { $0.rawValue == store.clock.speed }?.title ?? "\(Int(store.clock.speed))x").font(.app(.caption))
                    Spacer()
                    if store.clock.speed < TimeSpeed.lifeIn12h.rawValue {
                        Button("최대 속도로 살기") { store.setSpeed(.lifeIn12h) }.font(.app(.caption, .semibold))
                    }
                }
            }
        }
    }

    private func stat(_ t: String, _ v: String) -> some View {
        VStack(alignment: .leading, spacing: 2) {
            Text(t).font(.app(.caption2)).foregroundStyle(.tertiary)
            Text(v).font(.app(.caption, .semibold)).monospacedDigit()
        }
    }

    private var ladder: some View {
        Card {
            VStack(alignment: .leading, spacing: 10) {
                Text("재력 단계").font(.app(.headline, .semibold))
                ForEach(WealthTier.allCases) { t in
                    let unlocked = store.game.isUnlocked(t)
                    let current = t == store.currentTier
                    let threshold = GameState.threshold(for: t)
                    HStack(spacing: 12) {
                        Text(t.emoji).font(.app(.title3)).opacity(unlocked ? 1 : 0.4)
                        VStack(alignment: .leading, spacing: 2) {
                            Text(t.title).font(.app(.body, current ? .bold : .regular)).foregroundStyle(unlocked ? .primary : .secondary)
                            if !unlocked {
                                Text("순자산 \(Money.format(threshold, code: store.currency)) 이상일 때 승급 퀘스트").font(.app(.caption2)).foregroundStyle(.tertiary)
                            }
                        }
                        Spacer()
                        if current { Text("현재").font(.app(.caption, .semibold)).foregroundStyle(store.theme.accent) }
                        else if !unlocked { Image(systemName: "lock.fill").foregroundStyle(.tertiary) }
                        else { Image(systemName: "checkmark").foregroundStyle(.secondary) }
                    }
                    if GameState.nextTier(after: store.currentTier) == t, !unlocked, threshold > 0 {
                        let p = min(1, max(0, (store.netWorth / threshold).doubleValue))
                        GeometryReader { geo in
                            Capsule().fill(Color.primary.opacity(0.07)).overlay(alignment: .leading) { Capsule().fill(store.theme.accent).frame(width: geo.size.width * p) }
                        }.frame(height: 6)
                        Text("\(Int(p * 100))% · 다음 단계까지").font(.app(.caption2)).foregroundStyle(.secondary)
                    }
                }
            }
        }
    }

    @ViewBuilder private var flows: some View {
        if !store.game.flows.isEmpty || !store.game.outcomes.isEmpty {
            Card {
                VStack(alignment: .leading, spacing: 10) {
                    Text("선택이 만든 흐름").font(.app(.headline, .semibold))
                    ForEach(store.game.flows) { f in
                        HStack {
                            VStack(alignment: .leading, spacing: 2) {
                                Text(store.brand.resolve(f.counterparty)).font(.app(.subheadline, .semibold)).lineLimit(1)
                                Text("\(f.memo) · 매월 \(f.day)일" + (f.until.map { " · \($0.formatted(date: .abbreviated, time: .omitted))까지" } ?? "")).font(.app(.caption)).foregroundStyle(.secondary)
                            }
                            Spacer()
                            MoneyText(amount: f.amount, signed: true, font: .app(.subheadline, .semibold)).foregroundStyle(f.amount > 0 ? store.theme.accent : .primary)
                        }
                    }
                    ForEach(store.game.outcomes) { o in
                        HStack {
                            Image(systemName: "hourglass").foregroundStyle(.secondary)
                            Text(QuestEngine.shortTitle(o.questID)).font(.app(.subheadline)).lineLimit(1)
                            Spacer()
                            Text(o.fireDate.formatted(date: .abbreviated, time: .omitted)).font(.app(.caption)).foregroundStyle(.secondary)
                        }
                    }
                }
            }
        }
    }

    private var timeline: some View {
        Card {
            VStack(alignment: .leading, spacing: 10) {
                Text("선택의 기록").font(.app(.headline, .semibold))
                if store.game.answers.isEmpty { Text("아직 선택한 것이 없습니다. 며칠 안에 첫 제안이 옵니다.").font(.app(.subheadline)).foregroundStyle(.secondary) }
                ForEach(store.game.answers.reversed().prefix(80)) { a in
                    VStack(alignment: .leading, spacing: 3) {
                        HStack {
                            Text(a.date.formatted(date: .abbreviated, time: .omitted)).font(.app(.caption2)).foregroundStyle(.tertiary)
                            Spacer()
                            if a.cash != 0 { MoneyText(amount: a.cash, signed: true, font: .app(.caption, .semibold), hideable: false).foregroundStyle(a.cash > 0 ? store.theme.accent : .secondary) }
                        }
                        Text(a.title).font(.app(.subheadline, .semibold)).lineLimit(2)
                        Text("→ \(a.choice)").font(.app(.caption)).foregroundStyle(store.theme.accent)
                        if let o = a.outcome, !o.isEmpty { Text(o).font(.app(.caption)).foregroundStyle(.secondary) }
                    }
                    Divider()
                }
            }
        }
    }
}

/// 인생 결산 at the end of the lifespan.
struct EndingView: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    @State private var confirmRestart = false

    var body: some View {
        let tier = store.currentTier
        let best = store.game.answers.max { $0.cash < $1.cash }
        let worst = store.game.answers.min { $0.cash < $1.cash }
        NavigationStack {
            ScrollView {
                VStack(spacing: 18) {
                    Text(tier.emoji).font(.app(size: 64)).padding(.top, 20)
                    Text("\(store.game.lifespanYears)세, \(tier.title)(으)로 마쳤습니다").font(.app(.title2, .bold)).multilineTextAlignment(.center)
                    Card {
                        VStack(alignment: .leading, spacing: 8) {
                            row("최종 순자산", Money.format(store.netWorth, code: store.currency))
                            row("선택 횟수", "\(store.game.answers.count)번")
                            row("도달 단계", store.game.unlockedTiers.map(\.title).joined(separator: " → "))
                            if let b = best, b.cash > 0 { row("최고의 선택", "\(b.choice) (\(Money.format(b.cash, code: store.currency)))") }
                            if let w = worst, w.cash < 0 { row("뼈아픈 선택", "\(w.choice) (\(Money.format(w.cash, code: store.currency)))") }
                            row("무지출 챌린지", Money.format(store.challenge.totalSaved, code: store.currency))
                        }
                    }
                    Text(verdict(tier)).font(.app(.subheadline)).foregroundStyle(.secondary).multilineTextAlignment(.center).padding(.horizontal)
                    Button { confirmRestart = true } label: { Text("새 인생 시작").frame(maxWidth: .infinity).padding(.vertical, 6) }.buttonStyle(.borderedProminent)
                    Button("결과만 보고 닫기") { dismiss() }.font(.app(.subheadline))
                }
                .padding()
            }
            .background(store.theme.background)
            .navigationTitle("인생 결산")
            .navigationBarTitleDisplayMode(.inline)
            .confirmationDialog("모든 데이터를 지우고 처음부터 시작할까요?", isPresented: $confirmRestart, titleVisibility: .visible) {
                Button("새 인생 시작", role: .destructive) { store.restartLife(); dismiss() }
            }
        }
    }
    private func row(_ k: String, _ v: String) -> some View {
        HStack(alignment: .top) { Text(k).font(.app(.subheadline)).foregroundStyle(.secondary); Spacer(); Text(v).font(.app(.subheadline, .semibold)).multilineTextAlignment(.trailing) }
    }
    private func verdict(_ t: WealthTier) -> String {
        switch t {
        case .broke: return "빚으로 끝났습니다. 다음 생에는 리딩방과 대포통장을 멀리하세요."
        case .working: return "성실하게 살았습니다. 큰 모험은 없었지만 빚도 없습니다."
        case .middle: return "집도 있고 차도 있습니다. 평범함이 가장 어렵습니다."
        case .rich: return "자산가로 생을 마쳤습니다. 스톡옵션을 행사한 그날이 분기점이었습니다."
        case .tycoon: return "재벌로 끝났습니다. 12시간 만에 한 인생을 살아냈습니다."
        }
    }
}

import SwiftUI

extension WealthTier {
    /// Tier as the game core's level (0 거지 … 4 재벌).
    var level: TierLevel {
        switch self { case .broke: return 0; case .working: return 1; case .middle: return 2; case .rich: return 3; case .tycoon: return 4 }
    }
}

/// 알바 hub: the jobs open to the current tier, today's remaining shifts, and what each job has earned so far.
struct JobsHubView: View {
    @EnvironmentObject var store: BankStore

    var body: some View {
        ScrollView {
            VStack(spacing: 14) {
                Card {
                    VStack(alignment: .leading, spacing: 8) {
                        HStack {
                            Text("오늘의 알바").font(.app(.title3, .bold))
                            MockChip()
                            Spacer()
                            Text("\(store.currentTier.emoji) \(store.currentTier.title)").font(.app(.subheadline, .semibold)).foregroundStyle(.secondary)
                        }
                        Text("본업 말고 부업으로 생활비를 버는 미니게임이에요. 번 돈은 게임 속 계좌로 들어갑니다.")
                            .font(.app(.footnote)).foregroundStyle(.secondary)
                        HStack(spacing: 6) {
                            ForEach(0..<JobsState.shiftsPerDay, id: \.self) { i in
                                Image(systemName: i < store.shiftsLeftToday ? "bolt.fill" : "bolt")
                                    .foregroundStyle(i < store.shiftsLeftToday ? Color.yellow : Color.secondary)
                            }
                            Text(store.shiftsLeftToday > 0 ? "오늘 \(store.shiftsLeftToday)탕 더 뛸 수 있어요" : "오늘은 녹초… 내일 다시")
                                .font(.app(.caption, .medium)).foregroundStyle(.secondary)
                        }
                        .accessibilityIdentifier("jobs-energy")
                    }
                }
                ForEach(JobKind.allCases) { kind in
                    JobRow(kind: kind)
                }
            }
            .padding(16)
        }
        .background(store.theme.background.ignoresSafeArea())
        .navigationTitle("알바")
        .navigationBarTitleDisplayMode(.inline)
    }
}

private struct JobRow: View {
    @EnvironmentObject var store: BankStore
    let kind: JobKind

    var body: some View {
        let locked = kind.lockedReason(for: store.currentTier.level)
        let tired = store.shiftsLeftToday == 0
        NavigationLink(value: HomeRoute.job(kind)) {
            HStack(spacing: 14) {
                Text(kind.emoji).font(.system(size: 30))
                    .frame(width: 52, height: 52)
                    .background(RoundedRectangle(cornerRadius: 14, style: .continuous).fill(store.theme.accent.opacity(0.12)))
                VStack(alignment: .leading, spacing: 3) {
                    HStack(spacing: 6) {
                        Text(kind.title).font(.app(.headline, .semibold))
                        Text(kind.shiftLabel).font(.app(.caption2, .medium)).foregroundStyle(.secondary)
                    }
                    Text(locked ?? kind.pitch).font(.app(.caption)).foregroundStyle(.secondary).lineLimit(2)
                    if let best = store.jobs.best[kind.rawValue] {
                        Text("최고 \(best)점 · \(store.jobs.plays[kind.rawValue] ?? 0)회").font(.app(.caption2)).foregroundStyle(.tertiary)
                    }
                }
                Spacer()
                if locked != nil { Image(systemName: "lock.fill").foregroundStyle(.secondary) }
                else if kind == .dayTrading { Text("투자금 필요").font(.app(.caption2, .semibold)).foregroundStyle(.orange) }
                else { Text(Money.format(wage(kind, .B), code: store.currency)).font(.app(.subheadline, .semibold)).monospacedDigit() }
            }
            .padding(14)
            .background(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous).fill(store.theme.card))
            .opacity(locked != nil || tired ? 0.55 : 1)
        }
        .buttonStyle(.plain)
        .disabled(locked != nil || tired)
        .accessibilityIdentifier("job-\(kind.rawValue)")
    }
}

/// One shift: intro → (stake for 단타) → the game → payslip. The money is booked once, when the shift ends.
struct JobSessionView: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    let kind: JobKind

    enum Stage: Equatable { case intro, playing(seed: UInt64), done(ShiftResult) }
    @State private var stage: Stage = .intro
    @State private var stake: Decimal = 0
    @State private var leverage: Double = 1

    var body: some View {
        Group {
            switch stage {
            case .intro: intro
            case .playing(let seed): board(seed: seed)
            case .done(let r): payslip(r)
            }
        }
        .background(store.theme.background.ignoresSafeArea())
        .navigationTitle(kind.title)
        .navigationBarTitleDisplayMode(.inline)
        .navigationBarBackButtonHidden(stageIsPlaying)
    }

    private var stageIsPlaying: Bool { if case .playing = stage { return true } else { return false } }

    private func finish(_ r: ShiftResult) {
        guard case .playing = stage else { return }
        store.recordShift(r)
        stage = .done(r)
    }

    @ViewBuilder private func board(seed: UInt64) -> some View {
        switch kind {
        case .convenience: ConvenienceBoard(game: ConvenienceGame(seed: seed), onFinish: finish)
        case .cafe: CafeBoard(game: CafeGame(seed: seed), onFinish: finish)
        case .parcel: ParcelBoard(game: ParcelGame(seed: seed), onFinish: finish)
        case .designatedDriver: DriverBoard(game: DriverGame(seed: seed), onFinish: finish)
        case .serving: ServingBoard(game: ServingGame(seed: seed), onFinish: finish)
        case .flyers: FlyerBoard(game: FlyerGame(seed: seed), onFinish: finish)
        case .tutoring: TutoringBoard(game: TutoringGame(seed: seed), onFinish: finish)
        case .dayTrading: TradingBoard(game: DayTradingGame(seed: seed, stake: stake, leverage: leverage), onFinish: finish)
        }
    }

    private var stakes: [Decimal] { DayTradingGame.stakes(for: store.currentTier.level) }
    private var balance: Decimal { store.payAccount?.balance ?? 0 }

    private var intro: some View {
        ScrollView {
            VStack(spacing: 18) {
                Text(kind.emoji).font(.system(size: 72)).padding(.top, 24)
                Text(kind.title).font(.app(.title2, .bold))
                Text(kind.pitch).font(.app(.body)).foregroundStyle(.secondary).multilineTextAlignment(.center).padding(.horizontal, 24)
                Card {
                    VStack(alignment: .leading, spacing: 8) {
                        row("근무", kind.shiftLabel)
                        if kind == .dayTrading {
                            row("결과", "수익 또는 손실이 그대로 계좌에")
                        } else {
                            row("기본 급여 (B등급)", Money.format(wage(kind, .B), code: store.currency))
                            row("S등급", Money.format(wage(kind, .S), code: store.currency))
                        }
                        if kind == .designatedDriver { row("흠집 1곳", "수리비 −\(Money.format(DriverGame.repairPerScrape, code: store.currency))") }
                    }
                }
                .padding(.horizontal, 16)
                if kind == .dayTrading { stakePicker.padding(.horizontal, 16) }
                Button {
                    stage = .playing(seed: UInt64.random(in: 1...UInt64.max))
                } label: {
                    Text(kind == .dayTrading ? "장 시작" : "출근하기").font(.app(.headline, .bold)).frame(maxWidth: .infinity).padding(.vertical, 16)
                        .background(RoundedRectangle(cornerRadius: 16, style: .continuous).fill(store.theme.accent))
                        .foregroundStyle(store.theme.onAccent)
                }
                .disabled(kind == .dayTrading && (stake <= 0 || stake > balance))
                .padding(.horizontal, 16)
                .accessibilityIdentifier("job-start")
            }
            .padding(.bottom, 32)
        }
        .onAppear { if stake == 0 { stake = stakes.first(where: { $0 <= balance }) ?? 0 } }
    }

    private var stakePicker: some View {
        Card {
            VStack(alignment: .leading, spacing: 10) {
                Text("투자금").font(.app(.subheadline, .semibold))
                HStack {
                    ForEach(stakes, id: \.self) { s in
                        Button { stake = s } label: {
                            Text(Money.format(s, code: store.currency)).font(.app(.caption, .semibold)).padding(.vertical, 8).padding(.horizontal, 10)
                                .background(Capsule().fill(stake == s ? store.theme.accent : Color.secondary.opacity(0.12)))
                                .foregroundStyle(stake == s ? store.theme.onAccent : Color.primary)
                        }
                        .disabled(s > balance)
                    }
                }
                if DayTradingGame.leverageAllowed(for: store.currentTier.level) {
                    Toggle("신용 3배 (−30%면 반대매매)", isOn: Binding(get: { leverage > 1 }, set: { leverage = $0 ? 3 : 1 }))
                        .font(.app(.caption))
                }
                Text("계좌 잔액 \(Money.format(balance, code: store.currency)) · 잃어도 투자금까지만 잃어요").font(.app(.caption2)).foregroundStyle(.secondary)
            }
        }
    }

    private func row(_ k: String, _ v: String) -> some View {
        HStack { Text(k).font(.app(.subheadline)).foregroundStyle(.secondary); Spacer(); Text(v).font(.app(.subheadline, .semibold)) }
    }

    private func payslip(_ r: ShiftResult) -> some View {
        ScrollView {
            VStack(spacing: 16) {
                Text(r.grade.rawValue).font(.system(size: 80, weight: .black, design: .rounded))
                    .foregroundStyle(r.grade >= .A ? Color.orange : (r.grade == .F ? Color.red : store.theme.accent))
                    .padding(.top, 24)
                    .accessibilityIdentifier("job-grade")
                Text(r.grade.comment).font(.app(.headline))
                Card {
                    VStack(alignment: .leading, spacing: 8) {
                        ForEach(r.lines, id: \.self) { Text($0).font(.app(.subheadline)) }
                        Divider()
                        HStack {
                            Text(r.payout >= 0 ? "입금" : "출금").font(.app(.subheadline)).foregroundStyle(.secondary)
                            Spacer()
                            Text(Money.format(r.payout, code: store.currency, signed: true)).font(.app(.title3, .bold)).monospacedDigit()
                                .foregroundStyle(r.payout >= 0 ? Color.primary : Color.red)
                        }
                        Text("점수 \(r.score)").font(.app(.caption)).foregroundStyle(.secondary)
                    }
                }
                .padding(.horizontal, 16)
                Button { dismiss() } label: {
                    Text("확인").font(.app(.headline, .bold)).frame(maxWidth: .infinity).padding(.vertical, 16)
                        .background(RoundedRectangle(cornerRadius: 16, style: .continuous).fill(store.theme.accent))
                        .foregroundStyle(store.theme.onAccent)
                }
                .padding(.horizontal, 16)
                .accessibilityIdentifier("job-done")
            }
        }
    }
}

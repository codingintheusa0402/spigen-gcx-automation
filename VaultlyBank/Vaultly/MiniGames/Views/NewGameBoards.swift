import SwiftUI

// 배달 스쿠터 레이스 · 블랙잭 boards. Rules live in MiniGames/Core; these only draw and forward taps.

private func won(_ d: Decimal) -> String { Money.format(d, code: "KRW") }

// MARK: - 스쿠터 레이스

struct ScooterBoard: View {
    @State var game: ScooterGame
    let onFinish: (ShiftResult) -> Void

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            HStack {
                Text("🏠 \(game.delivered)건").font(.app(.caption, .semibold))
                Text("🍜 싣는 중 \(game.carrying)").font(.app(.caption))
                Spacer()
                Text("사고 \(game.crashes)/\(ScooterGame.maxCrashes)").font(.app(.caption2)).foregroundStyle(game.crashes > 0 ? .red : .secondary)
            }
            GeometryReader { g in
                let w = g.size.width, h = g.size.height
                ZStack {
                    Rectangle().fill(Color(red: 0.16, green: 0.17, blue: 0.2))
                    ForEach([-0.33, 0.33], id: \.self) { lx in
                        Path { p in
                            let x = w / 2 + lx * w / 2
                            var y = CGFloat(Int(game.distance * 6) % 40) - 40
                            while y < h { p.move(to: CGPoint(x: x, y: y)); p.addLine(to: CGPoint(x: x, y: y + 20)); y += 40 }
                        }
                        .stroke(Color.white.opacity(0.5), lineWidth: 3)
                    }
                    ForEach(game.markers, id: \.id) { m in
                        Text(m.kind.rawValue).font(.system(size: m.kind == .bus ? 40 : 32))
                            .position(x: w / 2 + m.x * w / 2 * 0.95, y: h * 0.85 - CGFloat(m.y) / 150 * h * 0.85)
                    }
                    Text("🛵").font(.system(size: 40)).rotationEffect(.degrees(game.vx * 6))
                        .position(x: w / 2 + game.x * w / 2 * 0.95, y: h * 0.85)
                    if game.carrying > 0 {
                        Text("\(game.carrying)").font(.system(.caption2, design: .rounded).weight(.black)).padding(4).background(Circle().fill(Color.orange)).foregroundStyle(.white)
                            .position(x: w / 2 + game.x * w / 2 * 0.95 + 18, y: h * 0.85 - 22)
                    }
                }
                .clipShape(RoundedRectangle(cornerRadius: 12))
            }
            .frame(height: 340)
            HStack(spacing: 10) {
                hold("◀", -1).accessibilityIdentifier("scooter-left")
                Text("부스트 🔥").font(.system(.subheadline, design: .rounded).weight(.bold))
                    .frame(maxWidth: .infinity).padding(.vertical, 22)
                    .background(RoundedRectangle(cornerRadius: 12).fill(game.boost ? Color.red : Color.secondary.opacity(0.2)))
                    .gesture(DragGesture(minimumDistance: 0).onChanged { _ in game.boost = true }.onEnded { _ in game.boost = false })
                    .accessibilityIdentifier("scooter-boost")
                hold("▶", 1).accessibilityIdentifier("scooter-right")
            }
            Text("🍜에서 픽업, 🏠에 배달 · 25초 넘기면 식어요 · 3번 넘어지면 응급실").font(.app(.caption2)).foregroundStyle(.secondary)
        }
    }

    private func hold(_ title: String, _ dir: Double) -> some View {
        Text(title).font(.system(.title3, design: .rounded).weight(.bold))
            .frame(maxWidth: .infinity).padding(.vertical, 22)
            .background(RoundedRectangle(cornerRadius: 12).fill(game.steer == dir ? Color.orange : Color.secondary.opacity(0.2)))
            .gesture(DragGesture(minimumDistance: 0).onChanged { _ in game.steer = dir }.onEnded { _ in game.steer = 0 })
    }
}

// MARK: - 블랙잭

struct CardBoard: View {
    @State var game: CardGame
    let onFinish: (ShiftResult) -> Void

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            VStack(spacing: 14) {
                hand(title: "딜러", cards: game.dealer, total: game.dealerHidden ? nil : game.dealerTotal, hideSecond: game.dealerHidden)
                hand(title: "나", cards: game.player, total: game.player.isEmpty ? nil : game.playerTotal, hideSecond: false)
            }
            .padding(14)
            .frame(maxWidth: .infinity)
            .background(RoundedRectangle(cornerRadius: 14).fill(Color(red: 0.1, green: 0.4, blue: 0.25)))
            HStack {
                VStack(alignment: .leading, spacing: 2) {
                    Text("칩").font(.app(.caption)).foregroundStyle(.secondary)
                    Text(won(game.chips)).font(.system(.title3, design: .monospaced).weight(.bold)).monospacedDigit()
                }
                Spacer()
                VStack(alignment: .trailing, spacing: 2) {
                    Text("손익").font(.app(.caption)).foregroundStyle(.secondary)
                    Text(Money.format(game.pnl, code: "KRW", signed: true)).font(.system(.subheadline, design: .monospaced).weight(.bold))
                        .foregroundStyle(game.pnl >= 0 ? Color.red : Color.blue)
                }
            }
            if game.phase == .betting {
                PixelButton(title: "딜 (베팅 \(won(game.bet)))", tint: game.chips >= game.bet ? .orange : .gray, big: true) { game.deal() }
                    .accessibilityIdentifier("card-deal")
            } else {
                HStack(spacing: 8) {
                    PixelButton(title: "히트", tint: game.phase == .player ? .blue : .gray) { game.hit() }.accessibilityIdentifier("card-hit")
                    PixelButton(title: "스탠드", tint: game.phase == .player ? .red : .gray) { game.stand() }.accessibilityIdentifier("card-stand")
                    PixelButton(title: "더블", tint: game.canDouble ? .purple : .gray) { game.double() }.accessibilityIdentifier("card-double")
                }
            }
            Text("게임 머니 전용 · 딜러는 17에서 스탠드 · 블랙잭 3:2").font(.app(.caption2)).foregroundStyle(.secondary)
        }
    }

    private func hand(title: String, cards: [CardGame.Card], total: Int?, hideSecond: Bool) -> some View {
        VStack(alignment: .leading, spacing: 6) {
            HStack {
                Text(title).font(.app(.caption, .semibold)).foregroundStyle(.white.opacity(0.8))
                if let t = total { Text("\(t)").font(.system(.caption, design: .monospaced).weight(.bold)).foregroundStyle(.white) }
            }
            HStack(spacing: 6) {
                ForEach(Array(cards.enumerated()), id: \.element.id) { i, c in
                    let hidden = hideSecond && i == 1
                    Text(hidden ? "🂠" : c.label)
                        .font(.system(size: 18, weight: .bold, design: .rounded))
                        .foregroundStyle(hidden ? Color.white : (c.red ? Color.red : Color.black))
                        .frame(width: 44, height: 60)
                        .background(RoundedRectangle(cornerRadius: 6).fill(hidden ? Color.blue : Color.white))
                        .overlay(RoundedRectangle(cornerRadius: 6).stroke(Color.black.opacity(0.2)))
                }
                if cards.isEmpty { Text("—").foregroundStyle(.white.opacity(0.5)).frame(height: 60) }
            }
        }
        .frame(maxWidth: .infinity, alignment: .leading)
    }
}

/// 만 18세 확인 — shown once before the first 블랙잭 table (game money only; still rated 17+/청소년이용불가).
struct AgeGateSheet: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    let onConfirm: () -> Void

    var body: some View {
        VStack(spacing: 16) {
            Text("🔞").font(.system(size: 56)).padding(.top, 24)
            Text("만 18세 이상만 이용할 수 있어요").font(.app(.title3, .bold))
            Text("블랙잭은 게임 머니로만 진행되는 오락이에요. 실제 돈·현금화·외부 결제와는 전혀 관계가 없지만, 사행성을 모사하므로 성인만 이용할 수 있습니다.")
                .font(.app(.subheadline)).foregroundStyle(.secondary).multilineTextAlignment(.center).padding(.horizontal, 24)
            Spacer()
            Button {
                store.update { $0.ageVerified = true }
                dismiss()
                onConfirm()
            } label: { Text("만 18세 이상입니다").frame(maxWidth: .infinity).padding(.vertical, 6) }
                .buttonStyle(.borderedProminent)
                .accessibilityIdentifier("age-confirm")
            Button("아니요") { dismiss() }.font(.app(.subheadline)).foregroundStyle(.secondary).accessibilityIdentifier("age-cancel")
        }
        .padding(24)
        .presentationDetents([.medium])
    }
}

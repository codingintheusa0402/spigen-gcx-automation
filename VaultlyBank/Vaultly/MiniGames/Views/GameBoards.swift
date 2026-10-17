import SwiftUI

// The eight job boards. Each owns its game state, draws it, forwards taps, and reports the result once.
// The rules live in MiniGames/Core (unit-tested); nothing here decides an outcome.

/// Shared frame: timer, status line, quit button, and the 30 fps game clock.
struct GameScaffold<Content: View>: View {
    @EnvironmentObject var store: BankStore
    let timeLeft: Double
    let status: String?
    let quit: () -> Void
    let tick: (Double) -> Void
    @ViewBuilder let content: () -> Content
    @State private var clock = Timer.publish(every: 1.0 / 30, on: .main, in: .common).autoconnect()
    @State private var last = Date()

    var body: some View {
        VStack(spacing: 10) {
            HStack {
                Label(String(format: "%d초", max(0, Int(timeLeft.rounded(.up)))), systemImage: "timer")
                    .font(.system(.headline, design: .monospaced).weight(.bold))
                    .foregroundStyle(timeLeft < 10 ? Color.red : Color.primary)
                Spacer()
                Button("그만하기", action: quit).font(.app(.subheadline, .semibold)).accessibilityIdentifier("job-quit")
            }
            Text(status ?? " ").font(.app(.footnote, .semibold)).foregroundStyle(.secondary)
                .frame(maxWidth: .infinity, alignment: .leading).lineLimit(1)
            content()
            Spacer(minLength: 0)
        }
        .padding(16)
        .onAppear { last = Date() }
        .onReceive(clock) { now in
            let dt = min(0.1, max(0, now.timeIntervalSince(last)))
            last = now
            tick(dt)
        }
    }
}

/// Chunky "retro" button used across the boards.
struct PixelButton: View {
    let title: String
    var tint: Color = .blue
    var big = false
    let action: () -> Void
    var body: some View {
        Button(action: action) {
            Text(title).font(.system(big ? .title3 : .subheadline, design: .rounded).weight(.bold))
                .frame(maxWidth: .infinity).padding(.vertical, big ? 16 : 10)
                .background(RoundedRectangle(cornerRadius: 10).fill(tint))
                .overlay(RoundedRectangle(cornerRadius: 10).stroke(Color.black.opacity(0.25), lineWidth: 2))
                .foregroundStyle(.white)
        }
        .buttonStyle(.plain)
    }
}

private func bar(_ ratio: Double, _ color: Color) -> some View {
    GeometryReader { g in
        ZStack(alignment: .leading) {
            Capsule().fill(Color.secondary.opacity(0.15))
            Capsule().fill(color).frame(width: g.size.width * max(0, min(1, ratio)))
        }
    }
    .frame(height: 6)
}

private func won(_ n: Int) -> String {
    let f = NumberFormatter(); f.numberStyle = .decimal
    return (f.string(from: NSNumber(value: n)) ?? "\(n)") + "원"
}

// MARK: - 편의점

struct ConvenienceBoard: View {
    @State var game: ConvenienceGame
    let onFinish: (ShiftResult) -> Void

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            HStack {
                Text("대기 줄").font(.app(.caption)).foregroundStyle(.secondary)
                Text(String(repeating: "🧍", count: max(0, game.queue.count - 1)))
                Spacer()
                Text("\(game.served)명 · 실수 \(game.mistakes)").font(.app(.caption, .semibold))
            }
            if let c = game.current {
                VStack(alignment: .leading, spacing: 8) {
                    Text("🧑 손님 \(c.id + 1)").font(.app(.headline))
                    LazyVGrid(columns: [GridItem(.adaptive(minimum: 96))], spacing: 8) {
                        ForEach(Array(c.items.enumerated()), id: \.offset) { i, item in
                            Button { game.scan(i) } label: {
                                VStack(spacing: 2) {
                                    Text(item.emoji).font(.system(size: 28))
                                    Text(item.name).font(.app(.caption2))
                                    Text(won(item.price)).font(.system(.caption2, design: .monospaced))
                                }
                                .frame(maxWidth: .infinity).padding(8)
                                .background(RoundedRectangle(cornerRadius: 10).fill(c.scanned[i] ? Color.green.opacity(0.2) : Color.secondary.opacity(0.1)))
                                .overlay(alignment: .topTrailing) { if c.scanned[i] { Image(systemName: "checkmark.circle.fill").foregroundStyle(.green).padding(4) } }
                            }
                            .buttonStyle(.plain)
                            .accessibilityIdentifier("conv-item-\(i)")
                        }
                    }
                    if game.phase == .change {
                        HStack {
                            VStack(alignment: .leading) {
                                Text("합계 \(won(c.total))").font(.system(.subheadline, design: .monospaced))
                                Text("받은 돈 \(won(c.paid))").font(.system(.subheadline, design: .monospaced))
                            }
                            Spacer()
                            VStack(alignment: .trailing) {
                                Text("거스름돈").font(.app(.caption)).foregroundStyle(.secondary)
                                Text(won(game.changeDrawer)).font(.system(.title2, design: .monospaced).weight(.bold))
                            }
                        }
                        LazyVGrid(columns: Array(repeating: GridItem(.flexible()), count: 3), spacing: 8) {
                            ForEach(ConvenienceGame.denominations, id: \.self) { d in
                                PixelButton(title: won(d), tint: d >= 1_000 ? Color.green : Color.gray) { game.addChange(d) }
                                    .accessibilityIdentifier("conv-cash-\(d)")
                            }
                        }
                        HStack {
                            PixelButton(title: "다시 세기", tint: Color.gray) { game.clearChange() }
                            PixelButton(title: "건네기", tint: Color.orange) { game.giveChange() }.accessibilityIdentifier("conv-give")
                        }
                    } else {
                        Text("상품을 하나씩 눌러 바코드를 찍으세요").font(.app(.caption)).foregroundStyle(.secondary)
                    }
                }
            }
        }
    }
}

// MARK: - 카페

struct CafeBoard: View {
    @State var game: CafeGame
    let onFinish: (ShiftResult) -> Void

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            Text("주문 (눌러서 지금 만든 잔 내주기)").font(.app(.caption)).foregroundStyle(.secondary).frame(maxWidth: .infinity, alignment: .leading)
            VStack(spacing: 6) {
                ForEach(Array(game.orders.enumerated()), id: \.element.id) { i, o in
                    Button { game.serve(to: o.id) } label: {
                        VStack(alignment: .leading, spacing: 4) {
                            Text("\(o.spec.drink.emoji) \(o.spec.label)").font(.app(.subheadline, .semibold))
                            bar(o.patience / o.maxPatience, o.patience < 6 ? .red : .green)
                        }
                        .padding(10).background(RoundedRectangle(cornerRadius: 10).fill(Color.secondary.opacity(0.1)))
                    }
                    .buttonStyle(.plain)
                    .accessibilityIdentifier("cafe-order-\(i)")
                }
            }
            Divider()
            Text("만드는 중: \(game.cup.drink.emoji) \(game.cup.label)").font(.app(.headline)).frame(maxWidth: .infinity, alignment: .leading)
            LazyVGrid(columns: Array(repeating: GridItem(.flexible()), count: 3), spacing: 6) {
                ForEach(CafeGame.Drink.allCases, id: \.self) { d in
                    PixelButton(title: "\(d.emoji) \(d.rawValue)", tint: game.cup.drink == d ? Color.brown : Color.gray) { game.choose(d) }
                        .accessibilityIdentifier("cafe-drink-\(d.rawValue)")
                }
            }
            HStack {
                PixelButton(title: "따뜻하게", tint: game.cup.iced ? Color.gray : Color.red) { game.setIced(false) }
                PixelButton(title: "아이스", tint: game.cup.iced ? Color.blue : Color.gray) { game.setIced(true) }
            }
            HStack {
                PixelButton(title: "톨", tint: game.cup.large ? Color.gray : Color.indigo) { game.setLarge(false) }
                PixelButton(title: "그란데", tint: game.cup.large ? Color.indigo : Color.gray) { game.setLarge(true) }
                PixelButton(title: game.cup.extraShot ? "샷 추가 ✓" : "샷 추가", tint: game.cup.extraShot ? Color.orange : Color.gray) { game.toggleShot() }
            }
        }
    }
}

// MARK: - 택배 상하차

struct ParcelBoard: View {
    @State var game: ParcelGame
    let onFinish: (ShiftResult) -> Void
    private let colors: [Color] = [.blue, .green, .orange]

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            HStack(alignment: .top) {
                VStack(alignment: .leading, spacing: 4) {
                    Text("지금 박스 → \(ParcelGame.regions[game.box.region])").font(.app(.headline)).foregroundStyle(colors[game.box.region])
                    boxShape(game.box, cell: 18)
                    bar(game.boxTimer / max(game.boxTime, 0.1), game.boxTimer < 1.5 ? .red : .orange).frame(width: 120)
                }
                Spacer()
                VStack(alignment: .trailing, spacing: 4) {
                    Text("다음").font(.app(.caption)).foregroundStyle(.secondary)
                    boxShape(game.nextBox, cell: 10).opacity(0.7)
                    Text("출발 \(game.departures)대").font(.app(.caption, .semibold))
                }
            }
            Text("트럭의 칸(열)을 눌러 박스를 내려놓으세요").font(.app(.caption)).foregroundStyle(.secondary)
            HStack(alignment: .bottom, spacing: 10) {
                ForEach(0..<3, id: \.self) { t in truck(t) }
            }
        }
    }

    private func boxShape(_ b: ParcelGame.Box, cell: CGFloat) -> some View {
        VStack(spacing: 1) {
            ForEach(0..<b.h, id: \.self) { _ in
                HStack(spacing: 1) {
                    ForEach(0..<b.w, id: \.self) { _ in Rectangle().fill(colors[b.region]).frame(width: cell, height: cell).overlay(Rectangle().stroke(Color.black.opacity(0.3))) }
                }
            }
        }
    }

    private func truck(_ t: Int) -> some View {
        let tr = game.trucks[t]
        return VStack(spacing: 4) {
            HStack(spacing: 1) {
                ForEach(0..<ParcelGame.width, id: \.self) { c in
                    VStack(spacing: 1) {
                        ForEach((0..<ParcelGame.height).reversed(), id: \.self) { r in
                            Rectangle().fill(tr.cells[r][c] ? colors[t].opacity(0.85) : Color.secondary.opacity(0.12))
                                .aspectRatio(1, contentMode: .fit)
                        }
                    }
                    .contentShape(Rectangle())
                    .onTapGesture { game.place(truck: t, col: min(c, ParcelGame.width - game.box.w)) }
                    .accessibilityIdentifier("parcel-t\(t)-c\(c)")
                }
            }
            .padding(3).background(RoundedRectangle(cornerRadius: 4).stroke(colors[t], lineWidth: 2))
            Text("🚚 \(ParcelGame.regions[t]) \(Int(tr.fill * 100))%").font(.app(.caption2, .semibold))
        }
    }
}

// MARK: - 대리운전

struct DriverBoard: View {
    @State var game: DriverGame
    let onFinish: (ShiftResult) -> Void

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            HStack {
                Text("🏠 \(Int(game.progress * 100))%").font(.app(.caption, .semibold))
                bar(game.progress, .green)
                Text("승차감 \(Int(game.comfort * 100)) · 흠집 \(game.scrapes)").font(.app(.caption2))
            }
            GeometryReader { g in
                let w = g.size.width, h = g.size.height
                ZStack {
                    Rectangle().fill(Color(white: 0.12))
                    ForEach([-0.33, 0.33], id: \.self) { lx in
                        Path { p in
                            let x = w / 2 + lx * w / 2
                            var y = CGFloat(Int(game.distance * 6) % 40) - 40
                            while y < h { p.move(to: CGPoint(x: x, y: y)); p.addLine(to: CGPoint(x: x, y: y + 20)); y += 40 }
                        }
                        .stroke(Color.yellow.opacity(0.7), lineWidth: 3)
                    }
                    ForEach(game.obstacles, id: \.id) { o in
                        Text(o.kind.rawValue).font(.system(size: 34))
                            .position(x: w / 2 + o.x * w / 2 * 0.95, y: h * 0.85 - CGFloat(o.y) / 140 * h * 0.85)
                    }
                    Text("🚘").font(.system(size: 40)).rotationEffect(.degrees(game.vx * 8))
                        .position(x: w / 2 + game.x * w / 2 * 0.95, y: h * 0.85)
                }
                .clipShape(RoundedRectangle(cornerRadius: 12))
            }
            .frame(height: 360)
            HStack(spacing: 12) {
                steerButton("◀ 왼쪽", -1).accessibilityIdentifier("drive-left")
                steerButton("오른쪽 ▶", 1).accessibilityIdentifier("drive-right")
            }
        }
    }

    private func steerButton(_ title: String, _ dir: Double) -> some View {
        Text(title).font(.system(.title3, design: .rounded).weight(.bold))
            .frame(maxWidth: .infinity).padding(.vertical, 22)
            .background(RoundedRectangle(cornerRadius: 12).fill(game.steer == dir ? Color.orange : Color.secondary.opacity(0.2)))
            .gesture(DragGesture(minimumDistance: 0).onChanged { _ in game.steer = dir }.onEnded { _ in game.steer = 0 })
    }
}

// MARK: - 서빙

struct ServingBoard: View {
    @State var game: ServingGame
    let onFinish: (ShiftResult) -> Void

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            LazyVGrid(columns: Array(repeating: GridItem(.flexible()), count: 3), spacing: 8) {
                ForEach(game.tables, id: \.number) { t in
                    Button { game.serve(table: t.number) } label: {
                        VStack(spacing: 4) {
                            Text("\(t.number)번").font(.app(.caption2, .semibold))
                            Text(t.wants.map { ServingGame.menu[$0].emoji } ?? "🪑").font(.system(size: 28))
                            if t.wants != nil { bar(t.wait / 30, t.wait < 8 ? .red : .green) } else { Color.clear.frame(height: 6) }
                        }
                        .padding(8).frame(maxWidth: .infinity)
                        .background(RoundedRectangle(cornerRadius: 10).fill(Color.brown.opacity(0.18)))
                    }
                    .buttonStyle(.plain)
                    .accessibilityIdentifier("serve-table-\(t.number)")
                }
            }
            Text("주방 (눌러서 쟁반에 올리기)").font(.app(.caption)).foregroundStyle(.secondary).frame(maxWidth: .infinity, alignment: .leading)
            HStack {
                ForEach(game.pass, id: \.id) { d in
                    Button { game.pickUp(d.id) } label: {
                        Text(ServingGame.menu[d.menu].emoji).font(.system(size: 30)).padding(6)
                            .background(Circle().fill(Color.secondary.opacity(0.12)))
                    }
                    .buttonStyle(.plain)
                    .accessibilityIdentifier("serve-pass-\(d.menu)")
                }
                Spacer()
            }
            .frame(height: 50)
            VStack(spacing: 6) {
                ZStack {
                    Capsule().fill(Color.gray.opacity(0.5)).frame(width: 220, height: 14)
                    HStack(spacing: 4) { ForEach(game.tray, id: \.id) { Text(ServingGame.menu[$0.menu].emoji).font(.system(size: 26)) } }
                        .offset(y: -18)
                }
                .rotationEffect(.degrees(game.tilt * 28))
                .frame(height: 60)
                Text(abs(game.tilt) > 0.6 ? "⚠️ 기울어요!" : "쟁반 균형").font(.app(.caption2, .semibold)).foregroundStyle(abs(game.tilt) > 0.6 ? Color.red : Color.secondary)
                HStack {
                    PixelButton(title: "◀ 잡기", tint: Color.gray) { game.balance(-1) }.accessibilityIdentifier("serve-left")
                    PixelButton(title: "잡기 ▶", tint: Color.gray) { game.balance(1) }.accessibilityIdentifier("serve-right")
                }
            }
        }
    }
}

// MARK: - 전단지

struct FlyerBoard: View {
    @State var game: FlyerGame
    let onFinish: (ShiftResult) -> Void

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            HStack {
                Text("📄 \(game.handed)/\(FlyerGame.quota)").font(.system(.title3, design: .monospaced).weight(.bold))
                Spacer()
                Text("완벽 \(game.perfect) · 째려봄 \(game.glares)").font(.app(.caption))
            }
            GeometryReader { g in
                let w = g.size.width
                ZStack {
                    RoundedRectangle(cornerRadius: 12).fill(Color.gray.opacity(0.15))
                    Rectangle().fill(Color.yellow.opacity(0.25))
                        .frame(width: w * FlyerGame.zoneHalf * 2).position(x: w * FlyerGame.zone, y: g.size.height / 2)
                    Rectangle().fill(Color.orange.opacity(0.35))
                        .frame(width: w * FlyerGame.perfectHalf * 2).position(x: w * FlyerGame.zone, y: g.size.height / 2)
                    ForEach(game.walkers.filter { !$0.done }, id: \.id) { p in
                        Button { game.hand(to: p.id) } label: {
                            VStack(spacing: 0) {
                                Text(p.mood.rawValue).font(.system(size: 16))
                                Text("🚶").font(.system(size: 38))
                            }
                        }
                        .buttonStyle(.plain)
                        .position(x: w * p.x, y: g.size.height / 2)
                    }
                    Text("🙋").font(.system(size: 36)).position(x: w * FlyerGame.zone, y: g.size.height - 24)
                }
            }
            .frame(height: 220)
            Text("노란 구역에 들어온 사람을 눌러 건네세요. 📱 통화 중인 사람은 피하세요.").font(.app(.caption)).foregroundStyle(.secondary)
            PixelButton(title: "가장 가까운 사람에게 건네기", tint: Color.orange, big: true) { game.hand() }.accessibilityIdentifier("flyer-hand")
        }
    }
}

// MARK: - 과외

struct TutoringBoard: View {
    @State var game: TutoringGame
    let onFinish: (ShiftResult) -> Void

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            if let q = game.current {
                HStack {
                    Text("\(game.index + 1)/\(game.questions.count) · \(q.subject)").font(.app(.caption, .semibold))
                    Spacer()
                    if game.streak >= 2 { Text("🔥 \(game.streak)").font(.app(.caption, .bold)) }
                }
                bar(game.questionTime / TutoringGame.perQuestion, game.questionTime < 3 ? .red : .blue)
                Text(q.prompt).font(.system(.title2, design: .rounded).weight(.bold)).multilineTextAlignment(.center)
                    .frame(maxWidth: .infinity, minHeight: 120)
                ForEach(Array(q.options.enumerated()), id: \.offset) { i, opt in
                    PixelButton(title: opt, tint: Color.indigo, big: true) { game.answer(i) }.accessibilityIdentifier("tutor-opt-\(i)")
                }
            }
        }
    }
}

// MARK: - 주식 단타

struct TradingBoard: View {
    @EnvironmentObject var store: BankStore
    @State var game: DayTradingGame
    let onFinish: (ShiftResult) -> Void

    var body: some View {
        GameScaffold(timeLeft: game.timeLeft, status: game.lastEvent, quit: { onFinish(game.result()) },
                     tick: { dt in game.tick(dt); if game.isOver { onFinish(game.result()) } }) {
            HStack(alignment: .firstTextBaseline) {
                Text(game.name).font(.app(.headline))
                Spacer()
                Text(won(Int(game.price))).font(.system(.title2, design: .monospaced).weight(.bold))
            }
            chart.frame(height: 220)
            HStack {
                VStack(alignment: .leading) {
                    Text("평가손익").font(.app(.caption)).foregroundStyle(.secondary)
                    Text(String(format: "%@%@", game.pnl >= 0 ? "+" : "−", won(Int(abs(game.pnl)))))
                        .font(.system(.title3, design: .monospaced).weight(.bold))
                        .foregroundStyle(game.pnl >= 0 ? Color.red : Color.blue)     // 한국식: 상승 빨강, 하락 파랑
                }
                Spacer()
                Text(game.holding ? "보유 \(Int(game.shares))주" : "현금 보유").font(.app(.caption, .semibold))
                if game.leverage > 1 { Text("신용 3배").font(.app(.caption2, .bold)).foregroundStyle(.orange) }
            }
            HStack {
                PixelButton(title: "매수", tint: game.holding ? .gray : .red, big: true) { game.buy() }.accessibilityIdentifier("trade-buy")
                PixelButton(title: "매도", tint: game.holding ? .blue : .gray, big: true) { game.sell() }.accessibilityIdentifier("trade-sell")
            }
            Text("투자금 \(Money.format(game.stake, code: store.currency)) · 끝나면 보유분은 자동 매도").font(.app(.caption2)).foregroundStyle(.secondary)
        }
    }

    private var chart: some View {
        GeometryReader { g in
            let h = game.history
            let lo = h.min() ?? 0, hi = max(h.max() ?? 1, lo + 1)
            Path { p in
                for (i, v) in h.enumerated() {
                    let pt = CGPoint(x: g.size.width * CGFloat(i) / CGFloat(max(1, h.count - 1)),
                                     y: g.size.height * (1 - CGFloat((v - lo) / (hi - lo))))
                    if i == 0 { p.move(to: pt) } else { p.addLine(to: pt) }
                }
            }
            .stroke((h.last ?? 0) >= (h.first ?? 0) ? Color.red : Color.blue, lineWidth: 2)
            .background(RoundedRectangle(cornerRadius: 10).fill(Color.secondary.opacity(0.08)))
        }
    }
}

import SwiftUI

/// Mock brokerage screen: simulated prices, dry-run orders, user-entered positions.
struct InvestView: View {
    @EnvironmentObject var store: BankStore
    enum SortMode: String, CaseIterable { case rateAsc = "수익률 낮은 순", rateDesc = "수익률 높은 순", valueDesc = "평가금 큰 순", name = "이름순" }
    @State private var sort: SortMode = .rateAsc
    @State private var showValue = true     // 평가금 vs 현재가
    @State private var showKRW = true       // 원 vs $
    @State private var adding = false

    private var p: Portfolio { store.portfolio }

    var body: some View {
        NavigationStack {
            ScrollView {
                VStack(alignment: .leading, spacing: 24) {
                    ticker
                    cashCards
                    summary
                    controls
                    marketSection(.domestic)
                    marketSection(.overseas)
                    Text("모의 투자입니다. 시세는 앱 안에서 무작위로 움직이며 실제 시장과 무관합니다.")
                        .font(.app(.caption2)).foregroundStyle(.tertiary)
                }
                .padding()
            }
            .background(store.theme.background)
            .navigationTitle(store.brand.broker)
            .navigationDestination(for: Holding.self) { HoldingDetailView(holdingID: $0.id) }
            .toolbar {
                ToolbarItem(placement: .topBarLeading) { SimBadge() }
                ToolbarItem(placement: .topBarTrailing) {
                    Button { adding = true } label: { Image(systemName: "plus") }
                }
            }
            .sheet(isPresented: $adding) { AddHoldingSheet() }
        }
    }

    private var ticker: some View {
        HStack(spacing: 8) {
            Text(p.indexName).font(.app(.subheadline, .semibold)).foregroundStyle(.secondary)
            Text(p.indexValue.formatted(.number.precision(.fractionLength(2)))).font(.app(.subheadline, .bold)).monospacedDigit()
            Text(String(format: "%@%.1f%%", p.indexChange >= 0 ? "+" : "", p.indexChange * 100))
                .font(.app(.caption, .semibold)).foregroundStyle(p.indexChange >= 0 ? .red : .blue)
            Spacer()
            if !p.simulatePrices { Text("시세 고정").font(.app(.caption2)).foregroundStyle(.tertiary) }
        }
    }

    private var cashCards: some View {
        HStack(spacing: 12) {
            cashCard("원화", Money.format(p.cashKRW, code: "KRW"))
            cashCard("달러", Money.format(p.cashUSD, code: "USD"))
        }
    }
    private func cashCard(_ title: String, _ value: String) -> some View {
        VStack(alignment: .leading, spacing: 6) {
            Text(title).font(.app(.subheadline)).foregroundStyle(.secondary)
            Text(store.settings.hideBalances ? "••••" : value).font(.app(.title3, .bold)).monospacedDigit()
        }
        .padding(18).frame(maxWidth: .infinity, alignment: .leading)
        .background(store.theme.card, in: RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
    }

    private var summary: some View {
        VStack(alignment: .leading, spacing: 6) {
            Text("내 투자").font(.app(.headline, .semibold))
            HStack(spacing: 6) {
                MoneyText(amount: p.investedKRW, font: .system(size: 38, weight: .heavy))
                Image(systemName: "chevron.right").foregroundStyle(.secondary)
            }
            PnLText(pnl: p.pnlKRW, rate: p.pnlRate, code: "KRW", font: .title3.weight(.semibold))
        }
    }

    private var controls: some View {
        HStack {
            Menu {
                ForEach(SortMode.allCases, id: \.self) { m in Button(m.rawValue) { sort = m } }
            } label: {
                HStack(spacing: 4) { Text(sort.rawValue); Image(systemName: "chevron.down") }
                    .font(.app(.subheadline)).padding(.horizontal, 12).padding(.vertical, 8)
                    .background(Color.primary.opacity(0.08), in: RoundedRectangle(cornerRadius: 10))
            }.foregroundStyle(.primary)
            Spacer()
            segment(["현재가", "평가금"], selected: showValue ? 1 : 0) { showValue = $0 == 1 }
            segment(["$", "원"], selected: showKRW ? 1 : 0) { showKRW = $0 == 1 }
        }
    }
    private func segment(_ items: [String], selected: Int, action: @escaping (Int) -> Void) -> some View {
        HStack(spacing: 0) {
            ForEach(items.indices, id: \.self) { i in
                Text(items[i]).font(.app(.subheadline, i == selected ? .bold : .regular))
                    .padding(.horizontal, 10).padding(.vertical, 6)
                    .background(i == selected ? Color.primary.opacity(0.15) : .clear, in: RoundedRectangle(cornerRadius: 8))
                    .foregroundStyle(i == selected ? .primary : .secondary)
                    .onTapGesture { action(i) }
            }
        }
        .padding(3).background(Color.primary.opacity(0.08), in: RoundedRectangle(cornerRadius: 10))
    }

    @ViewBuilder private func marketSection(_ market: Market) -> some View {
        let items = sorted(p.holdings.filter { $0.market == market })
        if !items.isEmpty {
            VStack(alignment: .leading, spacing: 4) {
                HStack(spacing: 8) {
                    Text(market.title).font(.app(.title3, .semibold))
                    Text(String(format: "%@%.1f%%", p.rate(for: market) >= 0 ? "+" : "", p.rate(for: market) * 100))
                        .font(.app(.title3, .semibold)).foregroundStyle(p.rate(for: market) >= 0 ? .red : .blue)
                    Image(systemName: "chevron.right").font(.app(.caption)).foregroundStyle(.secondary)
                }
                .padding(.bottom, 8)
                ForEach(items) { h in
                    NavigationLink(value: h) { HoldingRow(holding: h, showValue: showValue, showKRW: showKRW) }.buttonStyle(.plain)
                }
            }
        }
    }

    private func sorted(_ hs: [Holding]) -> [Holding] {
        switch sort {
        case .rateAsc: return hs.sorted { $0.pnlRate < $1.pnlRate }
        case .rateDesc: return hs.sorted { $0.pnlRate > $1.pnlRate }
        case .valueDesc: return hs.sorted { p.toKRW($0.value, market: $0.market) > p.toKRW($1.value, market: $1.market) }
        case .name: return hs.sorted { $0.name < $1.name }
        }
    }
}

extension Holding: Hashable {
    func hash(into hasher: inout Hasher) { hasher.combine(id) }
}

/// Korean-market convention: gains red, losses blue.
struct PnLText: View {
    @EnvironmentObject var store: BankStore
    let pnl: Decimal
    let rate: Double
    let code: String
    var font: Font = .subheadline
    var body: some View {
        let sign = pnl >= 0 ? "+" : "-"
        Text(store.settings.hideBalances ? "••••" : "\(sign)\(Money.format(abs(pnl), code: code)) (\(String(format: "%.1f", abs(rate) * 100))%)")
            .font(font).monospacedDigit()
            .foregroundStyle(pnl >= 0 ? Color.red : Color.blue)
    }
}

struct HoldingRow: View {
    @EnvironmentObject var store: BankStore
    let holding: Holding
    let showValue: Bool
    let showKRW: Bool

    var body: some View {
        let p = store.portfolio
        let code = showKRW ? "KRW" : holding.market.currency
        let convert: (Decimal) -> Decimal = { showKRW ? p.toKRW($0, market: holding.market) : $0 }
        HStack(spacing: 14) {
            ZStack {
                Circle().fill(store.theme.gradient(for: holding.colorIndex)).frame(width: 46, height: 46)
                Text(String(holding.symbol.prefix(2))).font(.app(.caption, .black)).foregroundStyle(.white)
            }
            VStack(alignment: .leading, spacing: 4) {
                Text(holding.name).font(.app(.title3, .medium))
                Text("\(holding.shares.formatted())주").font(.app(.subheadline)).foregroundStyle(.secondary)
            }
            Spacer()
            VStack(alignment: .trailing, spacing: 4) {
                Text(store.settings.hideBalances ? "••••" : Money.format(convert(showValue ? holding.value : holding.price), code: code))
                    .font(.app(.title3, .bold)).monospacedDigit()
                PnLText(pnl: convert(holding.pnl), rate: holding.pnlRate, code: code)
            }
        }
        .padding(.vertical, 12)
    }
}

struct HoldingDetailView: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    let holdingID: UUID
    @State private var orderBuy = true
    @State private var orderShares: Decimal = 1
    @State private var orderPrice: Decimal = 0
    @State private var showOrder = false
    @State private var setPrice = false
    @State private var newPrice: Decimal = 0
    @State private var result: String?

    var body: some View {
        if let h = store.portfolio.holdings.first(where: { $0.id == holdingID }) {
            let code = h.market.currency
            ScrollView {
                VStack(alignment: .leading, spacing: 20) {
                    VStack(alignment: .leading, spacing: 6) {
                        Text("\(h.name) · \(h.symbol)").font(.app(.subheadline)).foregroundStyle(.secondary)
                        Text(Money.format(h.price, code: code)).font(.app(size: 36, .heavy)).monospacedDigit()
                            .contentTransition(.numericText())
                        PnLText(pnl: h.pnl, rate: h.pnlRate, code: code, font: .headline)
                    }
                    Card {
                        VStack(spacing: 10) {
                            kv("보유 수량", "\(h.shares.formatted())주")
                            kv("평균 단가", Money.format(h.avgCost, code: code))
                            kv("평가 금액", Money.format(h.value, code: code))
                            kv("투자 원금", Money.format(h.cost, code: code))
                            if h.market == .overseas { kv("원화 환산", Money.format(store.portfolio.toKRW(h.value, market: .overseas), code: "KRW")) }
                        }
                    }
                    HStack(spacing: 12) {
                        Button { orderBuy = true; orderPrice = h.price; showOrder = true } label: { Text("구매하기").frame(maxWidth: .infinity).font(.app(.body, .bold)) }
                            .buttonStyle(.borderedProminent).tint(.red)
                        Button { orderBuy = false; orderPrice = h.price; showOrder = true } label: { Text("판매하기").frame(maxWidth: .infinity).font(.app(.body, .bold)) }
                            .buttonStyle(.borderedProminent).tint(.blue)
                    }
                    Button { newPrice = h.price; setPrice = true } label: { Label("현재가 직접 설정", systemImage: "slider.horizontal.3").frame(maxWidth: .infinity) }
                        .buttonStyle(.bordered)
                    let trades = store.portfolio.trades.filter { $0.holdingID == holdingID }.sorted { $0.date > $1.date }
                    if !trades.isEmpty {
                        Card {
                            Text("체결 내역").font(.app(.headline, .semibold))
                            ForEach(trades) { t in
                                HStack {
                                    Text(t.isBuy ? "매수" : "매도").font(.app(.subheadline, .semibold)).foregroundStyle(t.isBuy ? .red : .blue)
                                    Text(t.date.formatted(date: .numeric, time: .shortened)).font(.app(.caption)).foregroundStyle(.secondary)
                                    Spacer()
                                    Text("\(t.shares.formatted())주 · \(Money.format(t.price, code: code))").font(.app(.subheadline)).monospacedDigit()
                                }
                                if t.id != trades.last?.id { Divider() }
                            }
                        }
                    }
                }.padding()
            }
            .background(store.theme.background)
            .navigationTitle(h.name)
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .topBarTrailing) {
                    Button(role: .destructive) { dismiss(); store.removeHolding(holdingID) } label: { Image(systemName: "trash") }
                }
            }
            .sheet(isPresented: $showOrder) {
                NavigationStack {
                    Form {
                        Section(orderBuy ? "매수 주문" : "매도 주문") {
                            Stepper("수량: \(orderShares.formatted())주", value: $orderShares, in: 1...100_000)
                            HStack { Text("가격"); Spacer(); TextField("가격", value: $orderPrice, format: .number).keyboardType(.decimalPad).multilineTextAlignment(.trailing) }
                            kv("주문 금액", Money.format((orderShares * orderPrice).rounded(scale: 2), code: code))
                            kv("예수금", Money.format(h.market == .domestic ? store.portfolio.cashKRW : store.portfolio.cashUSD, code: code))
                        }
                        Section {
                            Button {
                                let ok = store.trade(holdingID: holdingID, buy: orderBuy, shares: orderShares, price: orderPrice)
                                result = ok ? "\(orderBuy ? "매수" : "매도") 체결 완료 (모의)" : "주문 실패: 예수금 또는 수량이 부족합니다"
                                showOrder = false
                            } label: { Text(orderBuy ? "매수" : "매도").frame(maxWidth: .infinity).font(.app(.body, .bold)) }
                            .buttonStyle(.borderedProminent).tint(orderBuy ? .red : .blue).listRowInsets(EdgeInsets())
                        } footer: { Text("모의 주문입니다. 실제 거래가 발생하지 않습니다.") }
                    }
                    .navigationTitle(h.name)
                    .toolbar { ToolbarItem(placement: .cancellationAction) { Button("취소") { showOrder = false } } }
                }
                .presentationDetents([.medium])
            }
            .alert("현재가 설정", isPresented: $setPrice) {
                TextField("가격", value: $newPrice, format: .number).keyboardType(.decimalPad)
                Button("저장") { store.setPrice(newPrice, of: holdingID) }
                Button("취소", role: .cancel) {}
            }
            .alert(result ?? "", isPresented: Binding(get: { result != nil }, set: { if !$0 { result = nil } })) { Button("확인") {} }
        } else {
            ContentUnavailableView("종목을 찾을 수 없습니다", systemImage: "questionmark.folder")
        }
    }

    private func kv(_ k: String, _ v: String) -> some View {
        HStack { Text(k).foregroundStyle(.secondary); Spacer(); Text(v).fontWeight(.medium).monospacedDigit() }.font(.app(.subheadline))
    }
}

struct AddHoldingSheet: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    @State private var name = ""
    @State private var symbol = ""
    @State private var market: Market = .domestic
    @State private var shares: Decimal = 10
    @State private var avgCost: Decimal = 0
    @State private var price: Decimal = 0
    @State private var volatility = 2.0
    @State private var color = 0

    var body: some View {
        NavigationStack {
            Form {
                Section("종목") {
                    TextField("종목명 (예: 네이버)", text: $name)
                    TextField("심볼 (예: NVR)", text: $symbol).textInputAutocapitalization(.characters)
                    Picker("시장", selection: $market) { ForEach(Market.allCases) { Text($0.title).tag($0) } }.pickerStyle(.segmented)
                }
                Section("포지션 (\(market.currency))") {
                    Stepper("수량: \(shares.formatted())주", value: $shares, in: 1...1_000_000)
                    HStack { Text("평균 단가"); Spacer(); TextField("0", value: $avgCost, format: .number).keyboardType(.decimalPad).multilineTextAlignment(.trailing) }
                    HStack { Text("현재가"); Spacer(); TextField("0", value: $price, format: .number).keyboardType(.decimalPad).multilineTextAlignment(.trailing) }
                    HStack { Text("변동성"); Spacer(); Text(String(format: "±%.1f%%/시간", volatility)) }
                    Slider(value: $volatility, in: 0...10, step: 0.5)
                }
                Section("아이콘 색") {
                    HStack(spacing: 12) {
                        ForEach(0..<6, id: \.self) { i in
                            Circle().fill(store.theme.gradient(for: i)).frame(width: 28, height: 28)
                                .overlay(Circle().stroke(Color.primary, lineWidth: color == i ? 2 : 0))
                                .onTapGesture { color = i }
                        }
                    }
                }
            }
            .navigationTitle("종목 추가")
            .toolbar {
                ToolbarItem(placement: .cancellationAction) { Button("취소") { dismiss() } }
                ToolbarItem(placement: .confirmationAction) {
                    Button("추가") {
                        store.addHolding(Holding(symbol: symbol.isEmpty ? String(name.prefix(3)).uppercased() : symbol, name: name, market: market,
                                                 shares: shares, avgCost: avgCost, price: price == 0 ? avgCost : price,
                                                 volatility: volatility / 100, colorIndex: color))
                        dismiss()
                    }.disabled(name.isEmpty || avgCost <= 0)
                }
            }
        }
    }
}

struct CashEditorSheet: View {
    @EnvironmentObject var store: BankStore
    @Environment(\.dismiss) private var dismiss
    @State private var krw: Decimal = 0
    @State private var usd: Decimal = 0
    @State private var fx: Decimal = 0
    @State private var indexName = ""
    @State private var indexValue: Decimal = 0

    var body: some View {
        NavigationStack {
            Form {
                Section("예수금") {
                    HStack { Text("원화"); Spacer(); TextField("0", value: $krw, format: .number).keyboardType(.numberPad).multilineTextAlignment(.trailing) }
                    HStack { Text("달러"); Spacer(); TextField("0", value: $usd, format: .number).keyboardType(.decimalPad).multilineTextAlignment(.trailing) }
                }
                Section("환율 (1달러 = 원)") {
                    TextField("1380", value: $fx, format: .number).keyboardType(.numberPad)
                }
                Section("지수 표시 (가상)") {
                    TextField("지수 이름", text: $indexName)
                    TextField("지수 값", value: $indexValue, format: .number).keyboardType(.decimalPad)
                }
            }
            .navigationTitle("예수금 · 환율")
            .onAppear {
                let p = store.portfolio
                krw = p.cashKRW; usd = p.cashUSD; fx = p.fxRate; indexName = p.indexName; indexValue = p.indexValue
            }
            .toolbar {
                ToolbarItem(placement: .cancellationAction) { Button("취소") { dismiss() } }
                ToolbarItem(placement: .confirmationAction) {
                    Button("저장") {
                        store.updatePortfolio { $0.cashKRW = krw; $0.cashUSD = usd; $0.fxRate = max(fx, 1); $0.indexName = indexName; $0.indexValue = indexValue }
                        dismiss()
                    }
                }
            }
        }
    }
}

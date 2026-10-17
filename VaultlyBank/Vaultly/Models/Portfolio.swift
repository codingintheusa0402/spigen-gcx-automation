import Foundation

enum Market: String, Codable, CaseIterable, Identifiable {
    case domestic, overseas
    var id: String { rawValue }
    var title: String { self == .domestic ? "국내주식" : "해외주식" }
    var currency: String { self == .domestic ? "KRW" : "USD" }
}

/// A mock stock position. Prices are simulated locally (random walk) or set by the user.
struct Holding: Identifiable, Codable, Equatable {
    var id = UUID()
    var symbol: String          // "NVR", "SNDK"
    var name: String            // "네이버", "샌디스크"
    var market: Market
    var shares: Decimal
    var avgCost: Decimal        // in market currency
    var price: Decimal          // current, in market currency
    var volatility: Double = 0.02   // per simulated hour, ±
    var colorIndex: Int = 0

    var value: Decimal { (shares * price).rounded(scale: 2) }
    var cost: Decimal { (shares * avgCost).rounded(scale: 2) }
    var pnl: Decimal { value - cost }
    var pnlRate: Double { cost == 0 ? 0 : (pnl / cost).doubleValue }
}

struct TradeRecord: Identifiable, Codable, Equatable {
    var id = UUID()
    var holdingID: UUID
    var date: Date
    var isBuy: Bool
    var shares: Decimal
    var price: Decimal
}

struct Portfolio: Codable, Equatable {
    var cashKRW: Decimal = 200_000
    var cashUSD: Decimal = 11.24
    var fxRate: Decimal = 1_380          // KRW per USD (user adjustable)
    var holdings: [Holding] = []
    var trades: [TradeRecord] = []
    var lastTick: Date = Date()
    var indexName: String = "VXI 종합"
    var indexValue: Decimal = 2_641.52
    var indexChange: Double = 0.0042
    var simulatePrices: Bool = true

    func toKRW(_ amount: Decimal, market: Market) -> Decimal {
        market == .domestic ? amount : (amount * fxRate).rounded()
    }
    var investedKRW: Decimal { holdings.reduce(0) { $0 + toKRW($1.value, market: $1.market) } }
    var costKRW: Decimal { holdings.reduce(0) { $0 + toKRW($1.cost, market: $1.market) } }
    var pnlKRW: Decimal { investedKRW - costKRW }
    var pnlRate: Double { costKRW == 0 ? 0 : (pnlKRW / costKRW).doubleValue }
    var cashKRWTotal: Decimal { cashKRW + (cashUSD * fxRate).rounded() }

    func rate(for market: Market) -> Double {
        let hs = holdings.filter { $0.market == market }
        let c = hs.reduce(Decimal(0)) { $0 + $1.cost }
        let v = hs.reduce(Decimal(0)) { $0 + $1.value }
        return c == 0 ? 0 : ((v - c) / c).doubleValue
    }

    static func starter() -> Portfolio {
        var p = Portfolio()
        p.holdings = [
            Holding(symbol: "NVR", name: "네이버", market: .domestic, shares: 14, avgCost: 223_500, price: 171_200, colorIndex: 2),
            Holding(symbol: "SNDK", name: "샌디스크", market: .overseas, shares: 1, avgCost: 1_702, price: 1_296, colorIndex: 4),
            Holding(symbol: "SOXL", name: "SOXL", market: .overseas, shares: 1, avgCost: 131, price: 100.4, volatility: 0.05, colorIndex: 4),
            Holding(symbol: "PODD", name: "인슐릿", market: .overseas, shares: 3, avgCost: 141.2, price: 130.1, colorIndex: 0),
        ]
        return p
    }
}

/// Simulated market: advances prices one random step per elapsed hour.
enum MarketSimulator {
    static func tick(now: Date = Date(), portfolio: inout Portfolio) {
        guard portfolio.simulatePrices else { portfolio.lastTick = now; return }
        let hours = Int(now.timeIntervalSince(portfolio.lastTick) / 3600)
        guard hours > 0 else { return }
        let steps = min(hours, 24 * 14)
        for i in portfolio.holdings.indices {
            var price = portfolio.holdings[i].price.doubleValue
            let vol = portfolio.holdings[i].volatility
            for _ in 0..<steps { price *= 1 + Double.random(in: -vol...vol) }
            let scale = portfolio.holdings[i].market == .domestic ? 0 : 2
            if price.isFinite { portfolio.holdings[i].price = max(Decimal(double: price).rounded(scale: scale), 1) }
        }
        var idx = portfolio.indexValue.doubleValue
        var change = 0.0
        for _ in 0..<steps { let r = Double.random(in: -0.004...0.004); idx *= 1 + r; change += r }
        if idx.isFinite { portfolio.indexValue = Decimal(double: idx).rounded(scale: 2) }
        portfolio.indexChange = change
        portfolio.lastTick = now
    }
}

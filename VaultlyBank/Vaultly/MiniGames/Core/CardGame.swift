import Foundation

/// 블랙잭 (게임 머니) — a 하우스 table played with in-game money only. Buy in with chips from the account; each hand bets
/// a tenth of the buy-in; the session ends when the clock runs out or the chips can't cover a bet.
/// Rules: 6-deck shoe, dealer stands on all 17s, blackjack pays 3:2, double on any first two cards, no splits/insurance.
struct CardGame: ShiftGame {
    struct Card: Equatable, Identifiable {
        let rank: Int        // 1 = A … 13 = K
        let suit: Int        // 0 ♠ 1 ♥ 2 ♦ 3 ♣
        var id: Int { rank * 10 + suit + 1000 * (serial) }
        let serial: Int
        var label: String { ["A", "2", "3", "4", "5", "6", "7", "8", "9", "10", "J", "Q", "K"][rank - 1] + ["♠", "♥", "♦", "♣"][suit] }
        var red: Bool { suit == 1 || suit == 2 }
        var value: Int { min(10, rank) }
    }
    enum Phase: Equatable { case betting, player, dealer, settled(String) }

    let kind = JobKind.cards
    static let sessionSeconds: Double = 120
    let stake: Decimal                 // buy-in from the account
    let bet: Decimal                   // per hand
    private(set) var chips: Decimal
    private(set) var timeLeft = CardGame.sessionSeconds
    private(set) var phase: Phase = .betting
    private(set) var player: [Card] = []
    private(set) var dealer: [Card] = []
    private(set) var hands = 0
    private(set) var wins = 0
    private(set) var blackjacks = 0
    private(set) var doubled = false
    private(set) var lastEvent: String?
    private var shoe: [Card] = []
    private var serial = 0
    private var rng: GameRNG
    private var settleTimer: Double = 0

    init(seed: UInt64, stake: Decimal) {
        rng = GameRNG(seed: seed)
        self.stake = stake
        bet = roundWon(stake / 10)
        chips = stake
    }

    static func stakes(for level: TierLevel) -> [Decimal] {
        switch level {
        case 0: return []                                     // 거지는 테이블에 못 앉는다
        case 1: return [50_000, 100_000]
        case 2: return [500_000, 2_000_000]
        case 3: return [10_000_000, 50_000_000]
        default: return [100_000_000, 1_000_000_000]
        }
    }

    var isOver: Bool { timeLeft <= 0 || (phase == .betting && chips < bet) }
    var pnl: Decimal { chips - stake }

    static func total(_ cards: [Card]) -> Int {
        var t = cards.reduce(0) { $0 + $1.value }
        var aces = cards.filter { $0.rank == 1 }.count
        while aces > 0 && t + 10 <= 21 { t += 10; aces -= 1 }
        return t
    }
    var playerTotal: Int { Self.total(player) }
    var dealerTotal: Int { Self.total(dealer) }
    var canDouble: Bool { phase == .player && player.count == 2 && chips >= bet * 2 }
    var dealerHidden: Bool { phase == .player }

    mutating func tick(_ dt: Double) {
        guard timeLeft > 0 else { return }
        timeLeft -= dt
        if case .settled = phase { settleTimer -= dt; if settleTimer <= 0 { phase = .betting; if isOver { timeLeft = 0 } } }
        if timeLeft <= 0, phase == .player { stand() }           // clock ran out mid-hand: play it out
    }

    private mutating func draw() -> Card {
        if shoe.count < 20 {
            var cards: [Card] = []
            for _ in 0..<6 { for r in 1...13 { for s in 0..<4 { serial += 1; cards.append(Card(rank: r, suit: s, serial: serial)) } } }
            for i in stride(from: cards.count - 1, to: 0, by: -1) { cards.swapAt(i, rng.int(0...i)) }
            shoe = cards
        }
        return shoe.removeLast()
    }

    mutating func deal() {
        guard phase == .betting, chips >= bet, timeLeft > 0 else { return }
        chips -= bet
        hands += 1
        doubled = false
        player = [draw(), draw()]
        dealer = [draw(), draw()]
        phase = .player
        lastEvent = "베팅 \(bet)원"
        if playerTotal == 21 { settleHand() }
    }

    mutating func hit() {
        guard phase == .player else { return }
        player.append(draw())
        if playerTotal >= 21 { settleHand() }
    }

    mutating func double() {
        guard canDouble else { return }
        chips -= bet; doubled = true
        player.append(draw())
        settleHand()
    }

    mutating func stand() {
        guard phase == .player else { return }
        settleHand()
    }

    private mutating func settleHand() {
        phase = .dealer
        let pt = playerTotal
        let playerBJ = player.count == 2 && pt == 21
        if pt <= 21 { while dealerTotal < 17 { dealer.append(draw()) } }
        let dtl = dealerTotal
        let dealerBJ = dealer.count == 2 && dtl == 21
        let wager = doubled ? bet * 2 : bet
        var text: String
        if pt > 21 { text = "버스트 \(pt) — 베팅 잃음" }
        else if playerBJ && !dealerBJ { chips += wager + roundWon(wager * Decimal(double: 1.5)); wins += 1; blackjacks += 1; text = "블랙잭! 3:2 지급" }
        else if dealerBJ && !playerBJ { text = "딜러 블랙잭" }
        else if dtl > 21 { chips += wager * 2; wins += 1; text = "딜러 버스트 \(dtl) — 승리" }
        else if pt > dtl { chips += wager * 2; wins += 1; text = "\(pt) vs \(dtl) — 승리" }
        else if pt == dtl { chips += wager; text = "\(pt) vs \(dtl) — 푸시 (베팅 반환)" }
        else { text = "\(pt) vs \(dtl) — 패" }
        lastEvent = text
        phase = .settled(text)
        settleTimer = 1.6
    }

    func result() -> ShiftResult {
        let net = pnl
        let ratio = hands == 0 ? 0 : 0.5 + (net / max(stake, 1)).doubleValue
        let grade = Grade.from(max(0, min(1, ratio)))
        let score = max(0, wins * 100 + blackjacks * 150 + Int((net / 1000).doubleValue))
        var lines = ["\(hands)판 · 승 \(wins) · 블랙잭 \(blackjacks)", "칩 \(chips)원 (바이인 \(stake)원)"]
        lines.append(net >= 0 ? "하우스에서 따온 돈" : "하우스가 가져간 돈")
        return ShiftResult(kind: kind, score: score, grade: grade, payout: roundWon(net), lines: lines)
    }
}

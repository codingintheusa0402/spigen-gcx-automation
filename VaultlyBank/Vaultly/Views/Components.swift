import SwiftUI

/// Always-visible "simulation" marker (App Store Guideline 1.1.6 — no fake/trick functionality).
struct SimBadge: View {
    var body: some View {
        Text("SIM")
            .font(.system(size: 10, weight: .black, design: .rounded))
            .padding(.horizontal, 6).padding(.vertical, 3)
            .background(Capsule().fill(Color.orange.opacity(0.9)))
            .foregroundStyle(.white)
            .accessibilityLabel("시뮬레이션 모드")
    }
}

struct MoneyText: View {
    @EnvironmentObject var store: BankStore
    let amount: Decimal
    var signed = false
    var font: Font = .body
    var hideable = true

    var body: some View {
        Text(store.settings.hideBalances && hideable ? "••••••" : Money.format(amount, code: store.currency, signed: signed))
            .font(font)
            .monospacedDigit()
            .contentTransition(.numericText())
            .animation(.snappy, value: amount)
    }
}

struct Card<Content: View>: View {
    @EnvironmentObject var store: BankStore
    let content: Content
    init(@ViewBuilder content: () -> Content) { self.content = content() }
    var body: some View {
        content
            .padding(16)
            .background(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous).fill(store.theme.card))
            .overlay(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous)
                .stroke(Color.primary.opacity(0.06)))
    }
}

/// Gradient "bank card" tile for an account.
struct AccountTile: View {
    @EnvironmentObject var store: BankStore
    let account: Account

    var body: some View {
        VStack(alignment: .leading, spacing: 10) {
            HStack {
                Image(systemName: account.kind.symbol)
                Text(account.kind.title).font(.caption.weight(.semibold))
                Spacer()
                if account.kind == .loan { Text("부채").font(.caption2.weight(.bold)).padding(4).background(.white.opacity(0.2), in: Capsule()) }
            }
            .foregroundStyle(.white.opacity(0.85))
            Text(account.name).font(.headline).foregroundStyle(.white).lineLimit(1)
            Text(account.number).font(.caption.monospaced()).foregroundStyle(.white.opacity(0.7))
            Spacer(minLength: 4)
            MoneyText(amount: account.kind == .loan ? -account.balance : account.balance,
                      font: .title3.weight(.bold))
                .foregroundStyle(.white)
            if let t = account.terms, !t.isClosed {
                Text(String(format: "연 %.2f%% · 만기 %@", t.annualRate * 100, t.maturityDate.formatted(date: .abbreviated, time: .omitted)))
                    .font(.caption2).foregroundStyle(.white.opacity(0.75))
            } else if account.terms?.isClosed == true {
                Text("종료된 상품").font(.caption2).foregroundStyle(.white.opacity(0.75))
            }
        }
        .padding(16)
        .frame(maxWidth: .infinity, minHeight: 150, alignment: .leading)
        .background(store.theme.gradient(for: account.colorIndex))
        .clipShape(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
        .shadow(color: .black.opacity(0.15), radius: 10, y: 5)
    }
}

struct TransactionRow: View {
    @EnvironmentObject var store: BankStore
    let tx: Transaction
    var showAccount = false

    var body: some View {
        HStack(spacing: 12) {
            ZStack {
                Circle().fill(store.theme.accent.opacity(0.12)).frame(width: 40, height: 40)
                Image(systemName: icon).foregroundStyle(store.theme.accent)
            }
            VStack(alignment: .leading, spacing: 2) {
                Text(tx.counterparty).font(.subheadline.weight(.semibold)).lineLimit(1)
                Text(subtitle).font(.caption).foregroundStyle(.secondary).lineLimit(1)
            }
            Spacer()
            VStack(alignment: .trailing, spacing: 2) {
                MoneyText(amount: tx.amount, signed: true, font: .subheadline.weight(.bold))
                    .foregroundStyle(tx.amount == 0 ? Color.secondary : (tx.amount > 0 ? Color.blue : Color.primary))
                MoneyText(amount: tx.balanceAfter, font: .caption2).foregroundStyle(.secondary)
            }
        }
        .padding(.vertical, 4)
    }

    private var icon: String {
        switch tx.category {
        case .transfer: return "arrow.left.arrow.right"
        case .scheduled: return "bell.badge.fill"
        case .interest: return "percent"
        case .autoTransfer: return "repeat"
        case .repayment: return "building.columns"
        case .adjustment: return "slider.horizontal.3"
        case .maturity: return "checkmark.seal.fill"
        }
    }

    private var subtitle: String {
        var parts = [tx.date.formatted(date: .numeric, time: .shortened), tx.memo]
        if showAccount, let a = store.account(tx.accountID) { parts.append(a.name) }
        return parts.filter { !$0.isEmpty }.joined(separator: " · ")
    }
}

/// Decimal text field that formats with grouping separators while typing.
struct AmountField: View {
    @EnvironmentObject var store: BankStore
    let title: String
    @Binding var amount: Decimal
    @State private var text = ""

    var body: some View {
        HStack {
            Text(title).foregroundStyle(.secondary)
            Spacer()
            TextField("0", text: $text)
                .keyboardType(store.currency == "KRW" ? .numberPad : .decimalPad)
                .multilineTextAlignment(.trailing)
                .font(.body.weight(.semibold).monospacedDigit())
                .onChange(of: text) { _, new in
                    let cleaned = new.filter { $0.isNumber || $0 == "." }
                    amount = Decimal(string: cleaned) ?? 0
                    let formatted = format(amount)
                    if formatted != new, !new.hasSuffix(".") { text = formatted }
                }
                .onAppear { if amount != 0 { text = format(amount) } }
                .onChange(of: amount) { _, a in
                    let typed = Decimal(string: text.filter { $0.isNumber || $0 == "." }) ?? 0
                    if typed != a { text = format(a) }
                }
            Text(store.currency == "KRW" ? "원" : store.currency).foregroundStyle(.secondary)
        }
    }

    private func format(_ d: Decimal) -> String {
        guard d != 0 else { return "" }
        let f = NumberFormatter(); f.numberStyle = .decimal; f.maximumFractionDigits = 2
        return f.string(from: NSDecimalNumber(decimal: d)) ?? ""
    }
}

struct BannerView: View {
    let message: BannerMessage
    var body: some View {
        HStack(spacing: 12) {
            Image(systemName: message.positive ? "arrow.down.circle.fill" : "arrow.up.circle.fill")
                .font(.title2).foregroundStyle(message.positive ? .green : .orange)
            VStack(alignment: .leading, spacing: 2) {
                Text(message.title).font(.subheadline.weight(.bold))
                Text(message.body).font(.caption).foregroundStyle(.secondary)
            }
            Spacer()
        }
        .padding(14)
        .background(.regularMaterial, in: RoundedRectangle(cornerRadius: 18, style: .continuous))
        .shadow(color: .black.opacity(0.15), radius: 12, y: 6)
        .padding(.horizontal)
    }
}

struct SuccessCheck: View {
    @State private var drawn = false
    var body: some View {
        ZStack {
            Circle().fill(Color.green.opacity(0.15)).frame(width: 120, height: 120)
            Circle().stroke(Color.green, lineWidth: 4).frame(width: 120, height: 120)
                .scaleEffect(drawn ? 1 : 0.6).opacity(drawn ? 1 : 0)
            Image(systemName: "checkmark").font(.system(size: 54, weight: .bold)).foregroundStyle(.green)
                .scaleEffect(drawn ? 1 : 0.2).opacity(drawn ? 1 : 0)
        }
        .onAppear { withAnimation(.spring(response: 0.5, dampingFraction: 0.6)) { drawn = true } }
    }
}

extension View {
    func themed(_ theme: Theme) -> some View {
        self.fontDesign(theme.fontDesign)
            .tint(theme.accent)
            .preferredColorScheme(theme.isDark ? .dark : nil)
    }
}

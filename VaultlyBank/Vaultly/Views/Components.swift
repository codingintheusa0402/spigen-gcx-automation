import SwiftUI

/// Always-visible "simulation" marker (App Store Guideline 1.1.6 — no fake/trick functionality).
struct SimBadge: View {
    var body: some View {
        Text("SIM")
            .font(.app(size: 10, .bold))
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
            .lineLimit(1)
            .minimumScaleFactor(0.55)
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

/// Flat account tile: colored icon, dark text on the card color — the way real bank apps draw accounts.
struct AccountTile: View {
    @EnvironmentObject var store: BankStore
    let account: Account

    var body: some View {
        HStack(alignment: .top, spacing: 14) {
            ZStack {
                Circle().fill(store.theme.gradient(for: account.colorIndex)).frame(width: 44, height: 44)
                Image(systemName: account.kind.symbol).font(.app(size: 17, .semibold)).foregroundStyle(.white)
            }
            VStack(alignment: .leading, spacing: 4) {
                HStack(spacing: 6) {
                    Text(account.kind.title).font(.app(.caption, .medium)).foregroundStyle(.secondary)
                    if account.kind == .loan {
                        Text("부채").font(.app(.caption2, .semibold)).padding(.horizontal, 6).padding(.vertical, 2)
                            .background(Color.primary.opacity(0.08), in: Capsule()).foregroundStyle(.secondary)
                    }
                }
                Text(account.name).font(.app(.body, .semibold)).lineLimit(1)
                Text(account.number).font(.app(.caption)).foregroundStyle(.secondary)
                MoneyText(amount: account.kind == .loan ? -account.balance : account.balance, font: .app(.title3, .bold))
                    .padding(.top, 6)
                if let t = account.terms, !t.isClosed {
                    Text(String(format: "연 %.2f%% · 만기 %@", t.annualRate * 100, t.maturityDate.formatted(date: .abbreviated, time: .omitted)))
                        .font(.app(.caption)).foregroundStyle(.secondary)
                } else if account.terms?.isClosed == true {
                    Text("종료된 상품").font(.app(.caption)).foregroundStyle(.secondary)
                }
            }
            Spacer(minLength: 0)
            Image(systemName: "chevron.right").font(.app(.caption)).foregroundStyle(.tertiary).padding(.top, 14)
        }
        .padding(16)
        .frame(maxWidth: .infinity, alignment: .leading)
        .background(store.theme.card)
        .clipShape(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous))
        .overlay(RoundedRectangle(cornerRadius: store.theme.cornerRadius, style: .continuous).stroke(Color.primary.opacity(0.06)))
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
                Text(tx.counterparty).font(.app(.subheadline, .semibold)).lineLimit(1)
                Text(subtitle).font(.app(.caption)).foregroundStyle(.secondary).lineLimit(1)
            }
            Spacer()
            VStack(alignment: .trailing, spacing: 2) {
                MoneyText(amount: tx.amount, signed: true, font: .app(.subheadline, .bold))
                    .foregroundStyle(tx.amount == 0 ? Color.secondary : (tx.amount > 0 ? Color.blue : Color.primary))
                MoneyText(amount: tx.balanceAfter, font: .app(.caption2)).foregroundStyle(.secondary)
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
                .font(.app(.body, .semibold).monospacedDigit())
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
                .font(.app(.title2)).foregroundStyle(message.positive ? .green : .orange)
            VStack(alignment: .leading, spacing: 2) {
                Text(message.title).font(.app(.subheadline, .bold))
                Text(message.body).font(.app(.caption)).foregroundStyle(.secondary)
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
            Image(systemName: "checkmark").font(.app(size: 54, .bold)).foregroundStyle(.green)
                .scaleEffect(drawn ? 1 : 0.2).opacity(drawn ? 1 : 0)
        }
        .onAppear { withAnimation(.spring(response: 0.5, dampingFraction: 0.6)) { drawn = true } }
    }
}

extension View {
    func themed(_ theme: Theme, appearance: Appearance) -> some View {
        let scheme: ColorScheme? = appearance == .light ? .light : appearance == .dark ? .dark : nil
        return self.font(.app(.body))
            .tint(theme.accent)
            .preferredColorScheme(scheme)
    }
}

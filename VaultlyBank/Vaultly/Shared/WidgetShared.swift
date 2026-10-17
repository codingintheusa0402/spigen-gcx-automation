import Foundation
import ActivityKit

/// Compiled into both the app and the widget extension.
enum WidgetShared {
    static let appGroup = "group.com.kevinkim.vaultly"
    static let snapshotFile = "widget-snapshot.json"
    static let kind = "VaultlyBalance"
    static let challengeKind = "VaultlyChallenge"

    static var containerURL: URL? {
        FileManager.default.containerURL(forSecurityApplicationGroupIdentifier: appGroup)
    }
}

/// Everything the widgets show. Written by the app after every settlement.
struct WidgetSnapshot: Codable {
    var bank: String = "Vaultly"
    var accountName: String = "입출금통장"
    var balance: Decimal = 0
    var totalAssets: Decimal = 0
    var wonSuffix: Bool = false
    var currency: String = "KRW"
    var hideBalances: Bool = false
    var streak: Int = 0
    var monthSaved: Decimal = 0
    var monthSpent: Decimal = 0
    var nextCounterparty: String? = nil
    var nextAmount: Decimal? = nil
    var nextDate: Date? = nil
    var updatedAt: Date = Date()

    static func load() -> WidgetSnapshot? {
        guard let url = WidgetShared.containerURL?.appendingPathComponent(WidgetShared.snapshotFile),
              let raw = try? Data(contentsOf: url) else { return nil }
        let dec = JSONDecoder(); dec.dateDecodingStrategy = .iso8601
        return try? dec.decode(WidgetSnapshot.self, from: raw)
    }

    func save() {
        guard let url = WidgetShared.containerURL?.appendingPathComponent(WidgetShared.snapshotFile) else { return }
        let enc = JSONEncoder(); enc.dateEncodingStrategy = .iso8601
        if let raw = try? enc.encode(self) { try? raw.write(to: url, options: .atomic) }
    }

    /// Formatting that does not depend on the app's Money global.
    func money(_ amount: Decimal) -> String {
        if hideBalances { return "••••••" }
        let f = NumberFormatter()
        if currency == "KRW" {
            f.numberStyle = .decimal; f.maximumFractionDigits = 0
            let body = f.string(from: NSDecimalNumber(decimal: abs(amount))) ?? "\(amount)"
            return wonSuffix ? "\(amount < 0 ? "-" : "")\(body)원" : "\(amount < 0 ? "-" : "")₩\(body)"
        }
        f.numberStyle = .currency; f.currencyCode = currency; f.locale = Locale(identifier: "en_US")
        return f.string(from: NSDecimalNumber(decimal: amount)) ?? "\(amount)"
    }
}

/// Live Activity: "입금 예정 16:00 · 엄마 5,000,000원".
struct DepositActivityAttributes: ActivityAttributes {
    struct ContentState: Codable, Hashable {
        var fireDate: Date
        var landed: Bool
    }
    var bank: String
    var counterparty: String
    var amountText: String
    var accountName: String
    var eventID: String
}

import Foundation
import StoreKit

/// StoreKit 2 wrapper: one lifetime unlock + one monthly subscription, both named Pro.
@MainActor
final class StoreService: ObservableObject {
    static let lifetimeID = "com.kevinkim.vaultly.pro"
    static let monthlyID = "com.kevinkim.vaultly.pro.monthly"

    @Published private(set) var products: [Product] = []
    @Published private(set) var purchased: Set<String> = []
    @Published var lastError: String?

    /// UI tests (`-uitest-pro`): behave as a Pro user in the Store build without StoreKit.
    static var debugPro = false

    /// Pro is always on outside the store variant.
    var isPro: Bool { !AppVariant.isStore || Self.debugPro || !purchased.isEmpty }

    private var updates: Task<Void, Never>?

    init() {
        updates = Task { [weak self] in
            for await result in StoreKit.Transaction.updates {
                if case .verified(let tx) = result {
                    await self?.refreshEntitlements()
                    await tx.finish()
                }
            }
        }
        Task { await load(); await refreshEntitlements() }
    }

    deinit { updates?.cancel() }

    func load() async {
        do { products = try await Product.products(for: [Self.lifetimeID, Self.monthlyID]).sorted { $0.price > $1.price } }
        catch { lastError = error.localizedDescription }
    }

    func refreshEntitlements() async {
        var ids: Set<String> = []
        for await result in StoreKit.Transaction.currentEntitlements {
            if case .verified(let tx) = result, tx.revocationDate == nil { ids.insert(tx.productID) }
        }
        purchased = ids
    }

    func purchase(_ product: Product) async {
        do {
            let result = try await product.purchase()
            if case .success(let verification) = result, case .verified(let tx) = verification {
                await tx.finish()
                await refreshEntitlements()
            }
        } catch { lastError = error.localizedDescription }
    }

    func restore() async {
        try? await AppStore.sync()
        await refreshEntitlements()
    }

    func allows(_ feature: ProFeature) -> Bool { isPro }
}

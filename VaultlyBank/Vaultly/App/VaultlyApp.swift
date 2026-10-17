import SwiftUI

@main
struct VaultlyApp: App {
    @UIApplicationDelegateAdaptor(AppDelegate.self) private var appDelegate
    @StateObject private var store: BankStore
    @StateObject private var shop = StoreService()
    @StateObject private var ads = AdService()
    @Environment(\.scenePhase) private var scenePhase

    init() {
        if CommandLine.arguments.contains("-uitest-reset") { Persistence.reset() }
        if CommandLine.arguments.contains("-uitest-pro") { StoreService.debugPro = true }
        AppFont.installUIKitAppearance()
        let store = BankStore()
        if let i = CommandLine.arguments.firstIndex(of: "-uitest-clock"), i + 1 < CommandLine.arguments.count {
            let f = DateFormatter(); f.dateFormat = "yyyy-MM-dd'T'HH:mm:ss"; f.locale = Locale(identifier: "en_US_POSIX")
            if let d = f.date(from: CommandLine.arguments[i + 1]) { store.uitestSetClock(d) }
        }
        if let i = CommandLine.arguments.firstIndex(of: "-uitest-theme"), i + 1 < CommandLine.arguments.count {
            let id = CommandLine.arguments[i + 1]
            store.update { $0.themeID = id }
        }
        if let i = CommandLine.arguments.firstIndex(of: "-uitest-quest"), i + 1 < CommandLine.arguments.count {
            store.update { $0.hasOnboarded = true }
            store.forcePendingQuest(CommandLine.arguments[i + 1])
        }
        if let i = CommandLine.arguments.firstIndex(of: "-uitest-tier"), i + 1 < CommandLine.arguments.count,
           let tier = WealthTier(rawValue: CommandLine.arguments[i + 1]) {
            store.uitestApplyTier(tier)
        }
        if let i = CommandLine.arguments.firstIndex(of: "-uitest-lifeevent"), i + 1 < CommandLine.arguments.count,
           let kind = LifeEventKind(rawValue: CommandLine.arguments[i + 1]) {
            store.uitestForceLifeEvent(kind)
        }
        _store = StateObject(wrappedValue: store)
        AppDelegate.store = store
    }

    var body: some Scene {
        WindowGroup {
            RootView()
                .environmentObject(store)
                .environmentObject(shop)
                .environmentObject(ads)
                .onAppear { let shop = self.shop; ads.isPro = { shop.isPro }; store.catchUp() }
                .onChange(of: scenePhase) { _, phase in
                    if phase == .active { store.catchUp() }
                }
        }
    }
}

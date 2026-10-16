import SwiftUI

@main
struct VaultlyApp: App {
    @StateObject private var store: BankStore
    @StateObject private var shop = StoreService()
    @Environment(\.scenePhase) private var scenePhase

    init() {
        if CommandLine.arguments.contains("-uitest-reset") { Persistence.reset() }
        AppFont.installUIKitAppearance()
        let store = BankStore()
        if let i = CommandLine.arguments.firstIndex(of: "-uitest-theme"), i + 1 < CommandLine.arguments.count {
            let id = CommandLine.arguments[i + 1]
            store.update { $0.themeID = id }
        }
        _store = StateObject(wrappedValue: store)
    }

    var body: some Scene {
        WindowGroup {
            RootView()
                .environmentObject(store)
                .environmentObject(shop)
                .onAppear { store.catchUp() }
                .onChange(of: scenePhase) { _, phase in
                    if phase == .active { store.catchUp() }
                }
        }
    }
}

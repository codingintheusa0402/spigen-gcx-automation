import SwiftUI

@main
struct VaultlyApp: App {
    @StateObject private var store: BankStore
    @Environment(\.scenePhase) private var scenePhase

    init() {
        if CommandLine.arguments.contains("-uitest-reset") { Persistence.reset() }
        AppFont.installUIKitAppearance()
        _store = StateObject(wrappedValue: BankStore())
    }

    var body: some Scene {
        WindowGroup {
            RootView()
                .environmentObject(store)
                .onAppear { store.catchUp() }
                .onChange(of: scenePhase) { _, phase in
                    if phase == .active { store.catchUp() }
                }
        }
    }
}

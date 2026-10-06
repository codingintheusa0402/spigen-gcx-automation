import SwiftUI

// Claude Mesh for iPhone — a native shell around the live mesh your Mac serves
// (Claude Mesh → 📱). The app remembers the paired address(es) and opens straight into the mesh.
@main
struct ClaudeMeshApp: App {
    @StateObject private var pairing = Pairing()
    var body: some Scene {
        WindowGroup {
            Group {
                if pairing.urls.isEmpty { PairingView() } else { MeshScreen() }
            }
            .environmentObject(pairing)
            .preferredColorScheme(.dark)
        }
    }
}

/// Saved pairing links, e.g. http://100.x.y.z:47320/?t=KEY (Tailscale) and http://192.168.x.y:47320/?t=KEY (Wi-Fi).
final class Pairing: ObservableObject {
    @Published var urls: [URL] { didSet { UserDefaults.standard.set(urls.map(\.absoluteString), forKey: "pairing.urls") } }
    init() { urls = (UserDefaults.standard.stringArray(forKey: "pairing.urls") ?? []).compactMap(URL.init(string:)) }

    /// Accepts the link from the Mac's QR code / copy. Returns false if it isn't a Claude Mesh pairing link.
    @discardableResult func add(_ text: String) -> Bool {
        guard let u = URL(string: text.trimmingCharacters(in: .whitespacesAndNewlines)),
              let scheme = u.scheme, scheme.hasPrefix("http"), u.host != nil, u.port == 47320,
              let q = URLComponents(url: u, resolvingAgainstBaseURL: false)?.queryItems,
              q.contains(where: { $0.name == "t" && !($0.value ?? "").isEmpty }) else { return false }
        // a new key replaces old links that carried a different key
        let key = q.first { $0.name == "t" }?.value
        urls.removeAll { URLComponents(url: $0, resolvingAgainstBaseURL: false)?.queryItems?.first { $0.name == "t" }?.value != key || $0.host == u.host }
        urls.insert(u, at: 0)
        return true
    }
    func reset() { urls = [] }
}

import SwiftUI
import WebKit

/// Full-screen live mesh. Tries each paired address in turn (e.g. Tailscale, then Wi-Fi);
/// if none answers, shows a native retry / re-pair screen.
struct MeshScreen: View {
    @EnvironmentObject var pairing: Pairing
    @State private var failed = false
    @State private var attempt = 0
    @State private var showMenu = false
    @Environment(\.scenePhase) private var phase

    var body: some View {
        ZStack {
            Color(red: 0.02, green: 0.016, blue: 0.10).ignoresSafeArea()
            MeshWebView(urls: pairing.urls, attempt: attempt, failed: $failed).ignoresSafeArea()
            if failed {
                VStack(spacing: 14) {
                    Image(systemName: "wifi.exclamationmark").font(.system(size: 40)).foregroundStyle(.white.opacity(0.7))
                    Text("Can’t reach your Mac").font(.headline).foregroundStyle(.white)
                    Text("Claude Mesh must be open on the Mac with phone access on.\nAway from home, turn on Tailscale on this iPhone.")
                        .font(.footnote).multilineTextAlignment(.center).foregroundStyle(.white.opacity(0.65))
                    HStack { Button("Retry") { failed = false; attempt += 1 }.buttonStyle(.borderedProminent)
                             Button("Re-pair") { pairing.reset() }.buttonStyle(.bordered) }
                }.padding(30).background(RoundedRectangle(cornerRadius: 24).fill(Color(red: 0.08, green: 0.07, blue: 0.25).opacity(0.95)))
            }
        }
        .onChange(of: phase) { _, p in if p == .active && failed { failed = false; attempt += 1 } }
        // long-press with three fingers anywhere → connection menu
        .overlay(ThreeFingerHold { showMenu = true }.allowsHitTesting(true).opacity(0.001))
        .confirmationDialog("Claude Mesh", isPresented: $showMenu) {
            Button("Reload") { attempt += 1 }
            Button("Re-pair with Mac", role: .destructive) { pairing.reset() }
        }
    }
}

struct MeshWebView: UIViewRepresentable {
    let urls: [URL]; let attempt: Int
    @Binding var failed: Bool

    func makeCoordinator() -> Coord { Coord(self) }
    func makeUIView(context: Context) -> WKWebView {
        let cfg = WKWebViewConfiguration()
        cfg.allowsInlineMediaPlayback = true
        cfg.websiteDataStore = .default()                       // keeps the pairing key + saved orbits
        let wv = WKWebView(frame: .zero, configuration: cfg)
        wv.isOpaque = false; wv.backgroundColor = UIColor(red: 0.02, green: 0.016, blue: 0.10, alpha: 1)
        wv.scrollView.bounces = false; wv.scrollView.contentInsetAdjustmentBehavior = .never
        wv.allowsBackForwardNavigationGestures = false
        wv.navigationDelegate = context.coordinator; wv.uiDelegate = context.coordinator
        context.coordinator.web = wv; context.coordinator.load(index: 0)
        return wv
    }
    func updateUIView(_ wv: WKWebView, context: Context) {
        if context.coordinator.lastAttempt != attempt { context.coordinator.lastAttempt = attempt; context.coordinator.load(index: 0) }
    }

    final class Coord: NSObject, WKNavigationDelegate, WKUIDelegate {
        var parent: MeshWebView; weak var web: WKWebView?; var index = 0; var lastAttempt = 0
        init(_ p: MeshWebView) { parent = p; lastAttempt = p.attempt }
        func load(index i: Int) {
            index = i
            guard i < parent.urls.count else { DispatchQueue.main.async { self.parent.failed = true }; return }
            var r = URLRequest(url: parent.urls[i]); r.timeoutInterval = 6
            web?.load(r)
        }
        // address didn't answer → try the next paired one
        func webView(_ w: WKWebView, didFailProvisionalNavigation n: WKNavigation!, withError e: Error) { load(index: index + 1) }
        func webView(_ w: WKWebView, didFail n: WKNavigation!, withError e: Error) { load(index: index + 1) }
        func webView(_ w: WKWebView, didFinish n: WKNavigation!) { DispatchQueue.main.async { self.parent.failed = false } }

        // the page's alert / confirm / prompt dialogs, shown natively
        private var top: UIViewController? { UIApplication.shared.connectedScenes.compactMap { ($0 as? UIWindowScene)?.keyWindow }.first?.rootViewController.flatMap { var v = $0; while let p = v.presentedViewController { v = p }; return v } }
        func webView(_ w: WKWebView, runJavaScriptAlertPanelWithMessage m: String, initiatedByFrame f: WKFrameInfo, completionHandler done: @escaping () -> Void) {
            let a = UIAlertController(title: nil, message: m, preferredStyle: .alert); a.addAction(.init(title: "OK", style: .default) { _ in done() })
            top?.present(a, animated: true) ?? done()
        }
        func webView(_ w: WKWebView, runJavaScriptConfirmPanelWithMessage m: String, initiatedByFrame f: WKFrameInfo, completionHandler done: @escaping (Bool) -> Void) {
            let a = UIAlertController(title: nil, message: m, preferredStyle: .alert)
            a.addAction(.init(title: "Cancel", style: .cancel) { _ in done(false) }); a.addAction(.init(title: "OK", style: .default) { _ in done(true) })
            top?.present(a, animated: true) ?? done(false)
        }
        func webView(_ w: WKWebView, runJavaScriptTextInputPanelWithPrompt p: String, defaultText: String?, initiatedByFrame f: WKFrameInfo, completionHandler done: @escaping (String?) -> Void) {
            let a = UIAlertController(title: nil, message: p, preferredStyle: .alert)
            a.addTextField { $0.text = defaultText }
            a.addAction(.init(title: "Cancel", style: .cancel) { _ in done(nil) }); a.addAction(.init(title: "OK", style: .default) { _ in done(a.textFields?.first?.text) })
            top?.present(a, animated: true) ?? done(nil)
        }
    }
}

/// Three-finger long press (doesn't interfere with the mesh's one/two-finger gestures).
struct ThreeFingerHold: UIViewRepresentable {
    var action: () -> Void
    func makeUIView(context: Context) -> UIView {
        let v = PassThroughView(); let g = UILongPressGestureRecognizer(target: context.coordinator, action: #selector(Coord.fire(_:)))
        g.numberOfTouchesRequired = 3; g.minimumPressDuration = 0.6; g.cancelsTouchesInView = false; v.addGestureRecognizer(g); return v
    }
    func updateUIView(_ uiView: UIView, context: Context) {}
    func makeCoordinator() -> Coord { Coord(action) }
    final class Coord: NSObject { let action: () -> Void; init(_ a: @escaping () -> Void) { action = a }
        @objc func fire(_ g: UILongPressGestureRecognizer) { if g.state == .began { action() } } }
    final class PassThroughView: UIView { override func point(inside p: CGPoint, with e: UIEvent?) -> Bool { (e?.allTouches?.count ?? 0) >= 3 } }
}

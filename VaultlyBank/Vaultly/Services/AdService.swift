import SwiftUI
import UIKit
#if STORE
import GoogleMobileAds
import UserMessagingPlatform
import AppTrackingTransparency
#endif

/// Ads exist only in the App Store build, and never for Pro. Rewarded ad = undo the last life choice;
/// interstitials after every few 알바 payslips / quest answers.
/// UI tests (any `-uitest-*` launch) get the in-process fake so the walkthroughs never touch the network.
@MainActor
final class AdService: ObservableObject {
    enum Mode { case off, fake, real }
    enum Placement { case afterShift, afterQuest }

    // Google's public TEST ids. Replace with the real AdMob app/unit ids before submitting (see AppStore/LISTING.md).
    static let appID = "ca-app-pub-3940256099942544~1458002511"
    static let rewardedUnit = "ca-app-pub-3940256099942544/1712485313"
    static let interstitialUnit = "ca-app-pub-3940256099942544/4411468910"
    static let shiftsPerInterstitial = 3
    static let questsPerInterstitial = 4

    struct FakeAd: Identifiable { let id = UUID(); let rewarded: Bool; let onDone: (Bool) -> Void }

    let mode: Mode
    @Published var fake: FakeAd?                 // the fake overlay, shown by RootView
    @Published private(set) var rewardedReady = false
    /// Who decides whether ads are off: StoreService.isPro, injected by the app.
    var isPro: () -> Bool = { true }
    private var shiftCount = 0
    private var questCount = 0
    private var started = false
    #if STORE
    private var rewarded: RewardedAd?
    private var interstitial: InterstitialAd?
    private var delegate: Dismissal?
    #endif

    init() {
        if !AppVariant.isStore { mode = .off }
        else if CommandLine.arguments.contains(where: { $0.hasPrefix("-uitest") }) { mode = .fake }
        else { mode = .real }
        if mode == .fake { rewardedReady = true }
    }

    var enabled: Bool { mode != .off && !isPro() }

    /// Call once the user is onboarded: tracking prompt (ATT), consent (UMP), SDK start, preload.
    func start() {
        guard enabled, !started else { return }
        started = true
        #if STORE
        if mode == .real { Task { await startReal() } }
        #endif
    }

    /// Rewarded: `completion(true)` only when the reward was actually earned.
    func showRewarded(_ completion: @escaping (Bool) -> Void) {
        guard enabled else { completion(true); return }      // Pro / personal build: the feature is free
        switch mode {
        case .off: completion(true)
        case .fake: fake = FakeAd(rewarded: true, onDone: completion)
        case .real:
            #if STORE
            presentRewarded(completion)
            #else
            completion(true)
            #endif
        }
    }

    /// Interstitial every Nth event of a placement; `completion` always runs (immediately when no ad is shown).
    func maybeInterstitial(_ placement: Placement, completion: @escaping () -> Void = {}) {
        guard enabled else { completion(); return }
        let due: Bool
        switch placement {
        case .afterShift: shiftCount += 1; due = shiftCount % Self.shiftsPerInterstitial == 0
        case .afterQuest: questCount += 1; due = questCount % Self.questsPerInterstitial == 0
        }
        guard due else { completion(); return }
        switch mode {
        case .off: completion()
        case .fake: fake = FakeAd(rewarded: false, onDone: { _ in completion() })
        case .real:
            #if STORE
            presentInterstitial(completion)
            #else
            completion()
            #endif
        }
    }

    static var rootViewController: UIViewController? {
        let scene = UIApplication.shared.connectedScenes.compactMap { $0 as? UIWindowScene }.first { $0.activationState == .foregroundActive }
            ?? UIApplication.shared.connectedScenes.compactMap { $0 as? UIWindowScene }.first
        var vc = scene?.windows.first { $0.isKeyWindow }?.rootViewController ?? scene?.windows.first?.rootViewController
        while let p = vc?.presentedViewController { vc = p }
        return vc
    }

    #if STORE
    private final class Dismissal: NSObject, FullScreenContentDelegate {
        var onDismiss: () -> Void = {}
        func adDidDismissFullScreenContent(_ ad: FullScreenPresentingAd) { onDismiss() }
        func ad(_ ad: FullScreenPresentingAd, didFailToPresentFullScreenContentWithError error: Error) { onDismiss() }
    }

    private func startReal() async {
        // 1. App Tracking Transparency — the system prompt, shown once (text: NSUserTrackingUsageDescription)
        if ATTrackingManager.trackingAuthorizationStatus == .notDetermined {
            _ = await ATTrackingManager.requestTrackingAuthorization()
        }
        // 2. GDPR/UMP consent where required (EEA/UK); elsewhere this returns immediately
        do {
            try await ConsentInformation.shared.requestConsentInfoUpdate(with: RequestParameters())
            if let vc = Self.rootViewController { try await ConsentForm.loadAndPresentIfRequired(from: vc) }
        } catch { }
        guard ConsentInformation.shared.canRequestAds else { return }
        // 3. SDK + preload
        await MobileAds.shared.start()
        await preload()
    }

    private func preload() async {
        if rewarded == nil { rewarded = try? await RewardedAd.load(with: Self.rewardedUnit, request: Request()); rewardedReady = rewarded != nil }
        if interstitial == nil { interstitial = try? await InterstitialAd.load(with: Self.interstitialUnit, request: Request()) }
    }

    private func presentRewarded(_ completion: @escaping (Bool) -> Void) {
        guard let ad = rewarded, let vc = Self.rootViewController else {
            Task { await preload() }
            completion(false); return
        }
        var earned = false
        let d = Dismissal()
        d.onDismiss = { [weak self] in
            self?.rewarded = nil; self?.rewardedReady = false
            completion(earned)
            Task { await self?.preload() }
        }
        delegate = d
        ad.fullScreenContentDelegate = d
        ad.present(from: vc) { earned = true }
    }

    private func presentInterstitial(_ completion: @escaping () -> Void) {
        guard let ad = interstitial, let vc = Self.rootViewController else { completion(); Task { await preload() }; return }
        let d = Dismissal()
        d.onDismiss = { [weak self] in self?.interstitial = nil; completion(); Task { await self?.preload() } }
        delegate = d
        ad.fullScreenContentDelegate = d
        ad.present(from: vc)
    }
    #endif
}

/// The fake ad used by UI tests and the simulator: a full-screen "광고" with a short countdown.
struct FakeAdOverlay: View {
    @EnvironmentObject var ads: AdService
    let ad: AdService.FakeAd
    @State private var left = 2

    var body: some View {
        ZStack {
            Color.black.opacity(0.92).ignoresSafeArea()
            VStack(spacing: 16) {
                Text("광고").font(.system(.caption, design: .rounded).weight(.bold)).foregroundStyle(.white.opacity(0.6))
                Text(ad.rewarded ? "🎁 보상형 광고 (테스트)" : "📺 전면 광고 (테스트)").font(.system(.title2, design: .rounded).weight(.bold)).foregroundStyle(.white)
                Text("실제 출시 빌드에서는 AdMob 광고가 여기에 나옵니다.").font(.footnote).foregroundStyle(.white.opacity(0.7))
                if left > 0 {
                    Text("\(left)초").font(.system(.title, design: .monospaced)).foregroundStyle(.white)
                } else if ad.rewarded {
                    Button { finish(true) } label: { Text("보상 받기").font(.headline).padding(.horizontal, 28).padding(.vertical, 12).background(Color.green, in: Capsule()).foregroundStyle(.white) }
                        .accessibilityIdentifier("ad-reward")
                    Button("건너뛰기 (보상 없음)") { finish(false) }.font(.footnote).foregroundStyle(.white.opacity(0.7)).accessibilityIdentifier("ad-skip")
                } else {
                    Button { finish(false) } label: { Text("닫기").font(.headline).padding(.horizontal, 28).padding(.vertical, 12).background(Color.white.opacity(0.2), in: Capsule()).foregroundStyle(.white) }
                        .accessibilityIdentifier("ad-close")
                }
            }
        }
        .accessibilityElement(children: .contain)
        .accessibilityIdentifier("ad-overlay")
        .task {
            while left > 0 { try? await Task.sleep(for: .seconds(1)); left -= 1 }
        }
    }

    private func finish(_ earned: Bool) {
        ads.fake = nil
        ad.onDone(earned)
    }
}

import UIKit
import UserNotifications

/// Routes quest notification taps/actions into the store.
final class AppDelegate: NSObject, UIApplicationDelegate, UNUserNotificationCenterDelegate {
    static weak var store: BankStore?

    func application(_ application: UIApplication, didFinishLaunchingWithOptions launchOptions: [UIApplication.LaunchOptionsKey: Any]? = nil) -> Bool {
        UNUserNotificationCenter.current().delegate = self
        NotificationService.registerCategories()
        return true
    }

    func userNotificationCenter(_ center: UNUserNotificationCenter, willPresent notification: UNNotification) async -> UNNotificationPresentationOptions {
        // The quest card is already on screen while the app is open; only alert outside the app.
        notification.request.content.categoryIdentifier == NotificationService.questCategory ? [] : [.banner, .sound, .list]
    }

    func userNotificationCenter(_ center: UNUserNotificationCenter, didReceive response: UNNotificationResponse) async {
        let info = response.notification.request.content.userInfo
        guard let questID = info["questID"] as? String else { return }
        let idx: Int? = response.actionIdentifier.hasPrefix("opt") ? Int(response.actionIdentifier.dropFirst(3)) : nil
        await MainActor.run { AppDelegate.store?.handleQuestNotification(questID: questID, optionIndex: idx) }
    }
}

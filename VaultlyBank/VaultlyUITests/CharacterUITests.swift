import XCTest

/// Pixel character on Home (every tier) and the LPC credits screen — screenshots for the layout check.
final class CharacterUITests: XCTestCase {
    private func snap(_ app: XCUIApplication, _ name: String) {
        let a = XCTAttachment(screenshot: app.screenshot()); a.name = name; a.lifetime = .keepAlways; add(a)
    }

    func testCharacterCardForEveryTier() {
        for tier in ["broke", "working", "middle", "rich", "tycoon"] {
            let app = XCUIApplication()
            app.launchArguments = ["-uitest-reset", "-uitest-tier", tier]
            app.launch()
            XCTAssertTrue(app.staticTexts["내 계좌"].waitForExistence(timeout: 5))
            XCTAssertTrue(app.descendants(matching: .any)["character-card"].firstMatch.waitForExistence(timeout: 5), "character card for \(tier)")
            XCTAssertTrue(app.descendants(matching: .any)["character-sprite"].firstMatch.exists, "sprite drawn for \(tier)")
            XCTAssertFalse(app.staticTexts["character-activity"].label.isEmpty)
            snap(app, "character-\(tier)")
            app.terminate()
        }
    }

    func testCreditsScreenListsLPCArtists() {
        let app = XCUIApplication()
        app.launchArguments = ["-uitest-reset", "-uitest-tier", "working"]
        app.launch()
        app.tabBars.firstMatch.buttons["설정"].tap()
        let link = app.buttons["settings-credits"]
        for _ in 0..<6 where !link.isHittable { app.swipeUp() }
        link.tap()
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS %@", "Liberated Pixel Cup")).firstMatch.waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS %@", "Stephen Challener")).firstMatch.exists, "LPC authors listed")
        snap(app, "credits")
    }
}

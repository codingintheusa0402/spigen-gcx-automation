import XCTest

/// 미니룸: the opening screen, one room per tier, furniture by wealth, clutter when broke, quick menu.
final class RoomUITests: XCTestCase {
    private var app: XCUIApplication!

    private func launch(_ extra: [String]) {
        continueAfterFailure = false
        app = XCUIApplication()
        app.launchArguments = ["-uitest-reset"] + extra
        app.launch()
    }
    private func snap(_ name: String) { let a = XCTAttachment(screenshot: app.screenshot()); a.name = name; a.lifetime = .keepAlways; add(a) }
    private func prop(_ name: String) -> XCUIElement { app.descendants(matching: .any)["room-prop-\(name)"].firstMatch }

    func testRoomReflectsWealthTier() {
        let rooms: [(String, String, String, String)] = [            // tier, always-there prop, prop of another tier, room name
            ("broke", "mattress", "sofa", "고시원"), ("working", "bed", "piano", "원룸"), ("middle", "sofa", "chandelier", "아파트"),
            ("rich", "window_city", "mattress", "펜트하우스"), ("tycoon", "chandelier", "cupramen", "저택")]
        for (tier, yes, no, name) in rooms {
            launch(["-uitest-tier", tier])
            XCTAssertTrue(app.otherElements["room-scene"].waitForExistence(timeout: 8), "room scene for \(tier)")
            XCTAssertTrue(prop(yes).exists, "\(tier) room has \(yes)")
            XCTAssertFalse(prop(no).exists, "\(tier) room must not show \(no)")
            XCTAssertTrue(app.staticTexts["room-name"].label.contains(name), "room name for \(tier): \(app.staticTexts["room-name"].label)")
            XCTAssertTrue(app.descendants(matching: .any)["character-sprite"].firstMatch.exists, "character lives in the room")
            XCTAssertTrue(app.staticTexts["room-status"].exists)
            snap("room-\(tier)")
            app.terminate()
        }
    }

    func testBrokeRoomShowsClutterAndWorkingRoomDoesNot() {
        launch(["-uitest-tier", "broke"])          // 거지 starts with a capital loan → in debt → bills and bottles
        XCTAssertTrue(app.otherElements["room-scene"].waitForExistence(timeout: 8))
        XCTAssertTrue(prop("bills").exists, "debt shows as a pile of bills")
        XCTAssertTrue(prop("soju").exists)
        app.terminate()
        launch(["-uitest-tier", "working"])
        XCTAssertTrue(app.otherElements["room-scene"].waitForExistence(timeout: 8))
        XCTAssertFalse(prop("bills").exists, "a funded account has no clutter")
        XCTAssertFalse(prop("piano").exists, "furniture past the current progress stays hidden")
    }

    func testOnboardingLetsYouStartAsBrokeOrWorking() {
        launch([])
        XCTAssertTrue(app.textFields["이름 (예: 홍길동)"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.buttons["start-working"].isSelected, "서민 is the default")
        app.buttons["start-broke"].tap()
        XCTAssertTrue(app.buttons["start-broke"].isSelected)
        snap("onboarding-start-choice")
        app.buttons["시작하기"].tap()
        let springboard = XCUIApplication(bundleIdentifier: "com.apple.springboard")
        let allow = springboard.buttons.matching(NSPredicate(format: "label IN %@", ["Allow", "허용"])).firstMatch
        if allow.waitForExistence(timeout: 3) { allow.tap() }
        XCTAssertTrue(app.otherElements["room-scene"].waitForExistence(timeout: 8))
        XCTAssertTrue(app.staticTexts["room-name"].label.contains("고시원"), "chose 거지 → 고시원")
        XCTAssertTrue(prop("mattress").exists)
        XCTAssertTrue(prop("bills").exists, "거지 starts in debt")
        snap("room-after-broke-start")
    }

    func testRoomIsTheOpeningScreenAndMenuNavigates() {
        launch([])
        let nameField = app.textFields["이름 (예: 홍길동)"]
        XCTAssertTrue(nameField.waitForExistence(timeout: 5))
        app.buttons["시작하기"].tap()
        let springboard = XCUIApplication(bundleIdentifier: "com.apple.springboard")
        let allow = springboard.buttons.matching(NSPredicate(format: "label IN %@", ["Allow", "허용"])).firstMatch
        if allow.waitForExistence(timeout: 3) { allow.tap() }
        XCTAssertTrue(app.navigationBars["미니룸"].waitForExistence(timeout: 8), "the app opens on the 미니룸")
        XCTAssertTrue(app.otherElements["room-scene"].exists)
        XCTAssertTrue(app.staticTexts["room-name"].label.contains("원룸"), "default start is 서민")
        snap("room-opening")
        app.buttons["room-menu-life"].tap()
        XCTAssertTrue(app.navigationBars["내 인생"].waitForExistence(timeout: 5))
        app.navigationBars.buttons.element(boundBy: 0).tap()
        app.buttons["room-menu-jobs"].tap()
        XCTAssertTrue(app.navigationBars["알바"].waitForExistence(timeout: 5))
        app.navigationBars.buttons.element(boundBy: 0).tap()
        app.buttons["room-menu-schedule"].tap()
        XCTAssertTrue(app.navigationBars["오늘의 일정"].waitForExistence(timeout: 5))
        app.navigationBars.buttons.element(boundBy: 0).tap()
        XCTAssertTrue(app.openBank(), "the bank home is one tab over")
        app.buttons["home-alerts"].tap()
        XCTAssertTrue(app.navigationBars["예약 알림"].waitForExistence(timeout: 5), "예약 알림 now lives under the bank home")
    }
}

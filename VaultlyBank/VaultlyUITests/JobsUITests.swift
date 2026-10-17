import XCTest

/// 알바 mini games: open every job, play a few real taps, quit, check the payslip, and screenshot each screen.
/// Runs on both the personal and the App Store scheme (the store build must show fictional employer names).
final class JobsUITests: XCTestCase {
    private var app: XCUIApplication!

    private func launch(_ extra: [String] = []) {
        continueAfterFailure = false
        app = XCUIApplication()
        app.launchArguments = ["-uitest-reset"] + extra
        app.launch()
    }

    private func snap(_ name: String) {
        let a = XCTAttachment(screenshot: app.screenshot())
        a.name = name
        a.lifetime = .keepAlways
        add(a)
    }

    private func onboard() {
        let nameField = app.textFields["이름 (예: 홍길동)"]
        XCTAssertTrue(nameField.waitForExistence(timeout: 5))
        nameField.tap(); nameField.typeText("Kevin")
        app.buttons["시작하기"].tap()
        let springboard = XCUIApplication(bundleIdentifier: "com.apple.springboard")
        let allow = springboard.buttons.matching(NSPredicate(format: "label IN %@", ["Allow", "허용"])).firstMatch
        if allow.waitForExistence(timeout: 3) { allow.tap() }
        XCTAssertTrue(app.openBank(timeout: 5))
    }

    private func openHub() {
        let tile = app.buttons["tile-jobs"]
        for _ in 0..<4 where !tile.isHittable { app.swipeUp() }
        XCTAssertTrue(tile.waitForExistence(timeout: 5), "알바 tile on Home")
        tile.tap()
        XCTAssertTrue(app.buttons["job-convenience"].waitForExistence(timeout: 5), "hub lists the jobs")
    }

    /// Start a job, run `play` for a moment, quit, verify the payslip, return to the hub.
    private func shift(_ kind: String, _ play: () -> Void) {
        let row = app.buttons["job-\(kind)"]
        for _ in 0..<4 where !row.isHittable { app.swipeUp() }
        XCTAssertTrue(row.isEnabled, "\(kind) should be available")
        row.tap()
        let start = app.buttons["job-start"]
        XCTAssertTrue(start.waitForExistence(timeout: 5))
        snap("job-\(kind)-intro")
        start.tap()
        XCTAssertTrue(app.buttons["job-quit"].waitForExistence(timeout: 5))
        play()
        snap("job-\(kind)-playing")
        app.buttons["job-quit"].tap()
        XCTAssertTrue(app.staticTexts["job-grade"].waitForExistence(timeout: 5) || app.otherElements["job-grade"].exists, "payslip for \(kind)")
        snap("job-\(kind)-payslip")
        app.buttons["job-done"].tap()
        if app.otherElements["ad-overlay"].waitForExistence(timeout: 1) {      // store build: interstitial every 3rd payslip
            let close = app.buttons["ad-close"]
            XCTAssertTrue(close.waitForExistence(timeout: 6)); close.tap()
        }
        XCTAssertTrue(app.buttons["job-convenience"].waitForExistence(timeout: 5), "back on the hub")
    }

    private func tapIfExists(_ id: String) { let b = app.buttons[id]; if b.exists && b.isHittable { b.tap() } }

    func testBrokeTierRetailJobsAndDailyLimit() {
        launch(); onboard(); openHub()
        snap("jobs-hub-broke")
        XCTAssertTrue(app.buttons["job-tutoring"].isEnabled, "every job is open to every tier")
        XCTAssertTrue(app.buttons["job-dayTrading"].isEnabled, "every job is open to every tier")

        shift("convenience") {
            for i in 0..<4 { tapIfExists("conv-item-\(i)") }
            tapIfExists("conv-cash-1000"); tapIfExists("conv-give")
            sleep(1)
        }
        shift("cafe") {
            sleep(3)
            tapIfExists("cafe-drink-카페라떼"); tapIfExists("cafe-order-0")
        }
        shift("parcel") {
            for c in 0..<3 { tapIfExists("parcel-t\(c)-c0"); usleep(300_000) }
        }
        // three shifts used: every job is now greyed out until tomorrow
        XCTAssertFalse(app.buttons["job-serving"].isEnabled, "daily energy is spent")
        snap("jobs-hub-tired")
    }

    func testBrokeTierPhysicalJobs() {
        launch(); onboard(); openHub()
        shift("designatedDriver") {
            for id in ["drive-left", "drive-right"] {
                let e = app.descendants(matching: .any)[id].firstMatch
                if e.waitForExistence(timeout: 2) { e.press(forDuration: 0.6) }
            }
            sleep(1)
        }
        shift("serving") {
            sleep(4)
            for m in 0..<8 { tapIfExists("serve-pass-\(m)") }
            tapIfExists("serve-left"); tapIfExists("serve-right")
            for t in 1...6 { tapIfExists("serve-table-\(t)") }
        }
        shift("flyers") {
            for _ in 0..<5 { tapIfExists("flyer-hand"); usleep(400_000) }
        }
    }

    func testMiddleTierTutoringAndTrading() {
        launch(["-uitest-tier", "middle"])
        XCTAssertTrue(app.openBank(timeout: 5))
        openHub()
        snap("jobs-hub-middle")
        XCTAssertTrue(app.buttons["job-parcel"].isEnabled, "every job is open to every tier")
        shift("tutoring") {
            for _ in 0..<3 { tapIfExists("tutor-opt-0"); usleep(300_000) }
        }
        shift("dayTrading") {
            tapIfExists("trade-buy"); sleep(2); tapIfExists("trade-sell")
        }
    }
}

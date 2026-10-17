import XCTest

/// 오늘의 일정(공휴일) · 인생 이벤트 · 되돌리기(보상형 광고) · 전면 광고 + Pro · 스쿠터 레이스 · 블랙잭(18+).
/// Runs on both schemes; the App Store scheme shows the in-process fake ads, the personal scheme has no ads.
final class LifeFeaturesUITests: XCTestCase {
    private var app: XCUIApplication!

    private func launch(_ extra: [String]) {
        continueAfterFailure = false
        app = XCUIApplication()
        app.launchArguments = ["-uitest-reset"] + extra
        app.launch()
        XCTAssertTrue(app.staticTexts["내 계좌"].waitForExistence(timeout: 8))
    }
    private func snap(_ name: String) {
        let a = XCTAttachment(screenshot: app.screenshot()); a.name = name; a.lifetime = .keepAlways; add(a)
    }
    private var isStoreBuild: Bool { app.staticTexts["모의 · 가상 자산"].firstMatch.exists }
    private func back() { app.navigationBars.buttons.element(boundBy: 0).tap() }
    private func scrollTo(_ e: XCUIElement) { for _ in 0..<6 where !e.isHittable { app.swipeUp() } }
    private func tapIfExists(_ id: String) { let b = app.buttons[id]; if b.exists && b.isHittable { b.tap() } }

    /// Dismiss a fake ad if one is up (store build). Returns whether one was shown.
    @discardableResult private func closeAdIfShown(reward: Bool) -> Bool {
        let overlay = app.otherElements["ad-overlay"]
        guard overlay.waitForExistence(timeout: 3) else { return false }
        let b = app.buttons[reward ? "ad-reward" : "ad-close"]
        XCTAssertTrue(b.waitForExistence(timeout: 6), "fake ad countdown ends")
        b.tap()
        return true
    }

    // MARK: 오늘의 일정 + 공휴일

    func testScheduleShowsHolidayAndTimetable() {
        launch(["-uitest-tier", "working", "-uitest-clock", "2026-10-03T10:00:00"])
        let card = app.descendants(matching: .any)["character-card"].firstMatch
        scrollTo(card)
        XCTAssertTrue(card.waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS '개천절'")).firstMatch.exists, "home card names today's holiday")
        app.buttons["character-schedule"].tap()
        XCTAssertTrue(app.navigationBars["오늘의 일정"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts["schedule-holiday"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts["schedule-holiday"].label.contains("개천절"))
        XCTAssertTrue(app.descendants(matching: .any)["schedule-list"].firstMatch.exists)
        XCTAssertTrue(app.staticTexts["schedule-next-holiday"].exists)
        snap("schedule-holiday")
        app.buttons["schedule-next"].tap()        // 10/4 Sunday
        app.buttons["schedule-next"].tap()        // 10/5 대체공휴일
        XCTAssertTrue(app.staticTexts["schedule-holiday"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts["schedule-holiday"].label.contains("대체공휴일"))
        snap("schedule-substitute")
        app.buttons["schedule-next"].tap()        // 10/6 Tuesday: a workday
        XCTAssertFalse(app.staticTexts["schedule-holiday"].exists)
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS '평일'")).firstMatch.exists)
        snap("schedule-workday")
    }

    // MARK: 인생 이벤트

    func testAccidentShowsInLifeTimelineAndClosesJobs() {
        launch(["-uitest-tier", "working", "-uitest-lifeevent", "accident"])
        app.buttons["life-row"].tap()
        XCTAssertTrue(app.navigationBars["내 인생"].waitForExistence(timeout: 5))
        let event = app.descendants(matching: .any)["life-event-accident"].firstMatch
        scrollTo(event)
        XCTAssertTrue(event.waitForExistence(timeout: 5), "교통사고 listed in the timeline")
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS '부상'")).firstMatch.exists)
        snap("life-event-accident")
        back()
        let tile = app.buttons["tile-jobs"]
        scrollTo(tile); tile.tap()
        XCTAssertTrue(app.staticTexts["jobs-closed"].waitForExistence(timeout: 5), "알바 closed while injured")
        XCTAssertFalse(app.buttons["job-convenience"].isEnabled)
        snap("jobs-closed-injured")
    }

    func testJailStopsSalaryAndShowsOnLifeScreen() {
        launch(["-uitest-tier", "working", "-uitest-lifeevent", "jail"])
        app.buttons["life-row"].tap()
        XCTAssertTrue(app.navigationBars["내 인생"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS '급여 중단'")).firstMatch.waitForExistence(timeout: 5))
        let event = app.descendants(matching: .any)["life-event-jail"].firstMatch
        scrollTo(event)
        XCTAssertTrue(event.exists)
        snap("life-event-jail")
    }

    // MARK: 되돌리기 (보상형 광고)

    func testUndoLastChoiceViaRewardedAd() {
        launch(["-uitest-quest", "ins_01"])
        XCTAssertTrue(app.buttons["quest-option-0"].waitForExistence(timeout: 5))
        app.buttons["quest-option-0"].tap()
        app.buttons["이렇게 한다"].tap()
        let undo = app.buttons["quest-undo"]
        XCTAssertTrue(undo.waitForExistence(timeout: 5), "outcome sheet offers 되돌리기")
        snap("outcome-undo")
        undo.tap()
        let shown = closeAdIfShown(reward: true)
        XCTAssertEqual(shown, isStoreBuild, "store build asks for a rewarded ad, personal build undoes for free")
        XCTAssertTrue(app.buttons["quest-option-0"].waitForExistence(timeout: 5), "the quest is back after undo")
        snap("quest-restored")
        // answer again: the second answer can also be undone, but only the latest one
        app.buttons["quest-option-1"].tap()
        app.buttons["이렇게 한다"].tap()
        XCTAssertTrue(app.buttons["확인"].waitForExistence(timeout: 5))
        app.buttons["확인"].tap()
        closeAdIfShown(reward: false)
        app.buttons["life-row"].tap()
        XCTAssertTrue(app.navigationBars["내 인생"].waitForExistence(timeout: 5))
        let timelineUndo = app.buttons["timeline-undo"]
        scrollTo(timelineUndo)
        XCTAssertTrue(timelineUndo.waitForExistence(timeout: 5), "latest answer in the timeline has 되돌리기")
    }

    // MARK: 전면 광고 + Pro

    func testInterstitialAfterShiftsAndProRemovesAds() {
        launch(["-uitest-tier", "working"])
        let tile = app.buttons["tile-jobs"]
        scrollTo(tile); tile.tap()
        var adSeen = false
        for i in 0..<3 {
            let row = app.buttons["job-flyers"]
            scrollTo(row); row.tap()
            XCTAssertTrue(app.buttons["job-start"].waitForExistence(timeout: 5)); app.buttons["job-start"].tap()
            XCTAssertTrue(app.buttons["job-quit"].waitForExistence(timeout: 5)); app.buttons["job-quit"].tap()
            XCTAssertTrue(app.buttons["job-done"].waitForExistence(timeout: 5)); app.buttons["job-done"].tap()
            if closeAdIfShown(reward: false) { adSeen = true; XCTAssertEqual(i, 2, "interstitial after the 3rd payslip") }
            XCTAssertTrue(app.buttons["job-convenience"].waitForExistence(timeout: 5))
        }
        XCTAssertEqual(adSeen, isStoreBuild)

        // Pro: no ads, bonus game money
        app.terminate()
        app.launchArguments = ["-uitest-reset", "-uitest-tier", "working", "-uitest-pro"]
        app.launch()
        XCTAssertTrue(app.staticTexts["내 계좌"].waitForExistence(timeout: 8))
        if isStoreBuild {
            XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS 'Pro 보너스'")).firstMatch.waitForExistence(timeout: 5), "Pro bonus deposit on home")
            snap("pro-bonus")
        }
        scrollTo(app.buttons["tile-jobs"]); app.buttons["tile-jobs"].tap()
        for _ in 0..<3 {
            let row = app.buttons["job-flyers"]
            scrollTo(row); row.tap()
            app.buttons["job-start"].tap()
            XCTAssertTrue(app.buttons["job-quit"].waitForExistence(timeout: 5)); app.buttons["job-quit"].tap()
            XCTAssertTrue(app.buttons["job-done"].waitForExistence(timeout: 5)); app.buttons["job-done"].tap()
            XCTAssertFalse(app.otherElements["ad-overlay"].waitForExistence(timeout: 1), "Pro never sees an ad")
        }
    }

    // MARK: 스쿠터 레이스 · 블랙잭

    func testScooterRaceAndBlackjack() {
        launch(["-uitest-tier", "working"])
        scrollTo(app.buttons["tile-jobs"]); app.buttons["tile-jobs"].tap()
        let scooter = app.buttons["job-scooter"]
        scrollTo(scooter)
        XCTAssertTrue(scooter.waitForExistence(timeout: 5)); scooter.tap()
        XCTAssertTrue(app.buttons["job-start"].waitForExistence(timeout: 5)); snap("scooter-intro"); app.buttons["job-start"].tap()
        XCTAssertTrue(app.buttons["job-quit"].waitForExistence(timeout: 5))
        for id in ["scooter-left", "scooter-right", "scooter-boost"] {
            let e = app.descendants(matching: .any)[id].firstMatch
            if e.waitForExistence(timeout: 2) { e.press(forDuration: 0.5) }
        }
        snap("scooter-playing")
        app.buttons["job-quit"].tap()
        XCTAssertTrue(app.staticTexts["job-grade"].waitForExistence(timeout: 5)); snap("scooter-payslip"); app.buttons["job-done"].tap()
        closeAdIfShown(reward: false)

        let cards = app.buttons["job-cards"]
        scrollTo(cards)
        XCTAssertTrue(cards.waitForExistence(timeout: 5), "블랙잭 listed under 오락")
        cards.tap()
        XCTAssertTrue(app.buttons["age-confirm"].waitForExistence(timeout: 5), "18+ gate before the first table")
        snap("cards-age-gate")
        app.buttons["age-confirm"].tap()
        XCTAssertTrue(app.buttons["job-start"].waitForExistence(timeout: 5)); snap("cards-intro"); app.buttons["job-start"].tap()
        XCTAssertTrue(app.buttons["card-deal"].waitForExistence(timeout: 5))
        app.buttons["card-deal"].tap()
        sleep(1)
        tapIfExists("card-hit"); tapIfExists("card-stand")
        snap("cards-playing")
        app.buttons["job-quit"].tap()
        XCTAssertTrue(app.staticTexts["job-grade"].waitForExistence(timeout: 5)); snap("cards-payslip"); app.buttons["job-done"].tap()
        closeAdIfShown(reward: false)
        XCTAssertTrue(app.buttons["job-convenience"].waitForExistence(timeout: 5))
        // 거지 can't sit at the table
        app.terminate()
        app.launchArguments = ["-uitest-reset", "-uitest-tier", "broke"]
        app.launch()
        XCTAssertTrue(app.staticTexts["내 계좌"].waitForExistence(timeout: 8))
        scrollTo(app.buttons["tile-jobs"]); app.buttons["tile-jobs"].tap()
        scrollTo(app.buttons["job-cards"])
        XCTAssertFalse(app.buttons["job-cards"].isEnabled, "블랙잭 locked for 거지")
    }
}

import XCTest

final class SmokeTests: XCTestCase {
    private var app: XCUIApplication!

    override func setUp() {
        continueAfterFailure = false
        app = XCUIApplication()
        app.launchArguments = ["-uitest-reset"]
        app.launch()
    }

    private func snap(_ name: String) {
        let a = XCTAttachment(screenshot: app.screenshot())
        a.name = name
        a.lifetime = .keepAlways
        add(a)
    }

    /// The App Store build (STORE flag) tiles a "모의 · 가상 자산" watermark over every screen and hides direct
    /// balance editing and hand-made scheduled deposits — the tests below branch on it.
    private var isStoreBuild: Bool { app.staticTexts["모의 · 가상 자산"].firstMatch.exists }

    private func allowNotificationsIfAsked() {
        let springboard = XCUIApplication(bundleIdentifier: "com.apple.springboard")
        let allow = springboard.buttons.matching(NSPredicate(format: "label IN %@", ["Allow", "허용"])).firstMatch
        if allow.waitForExistence(timeout: 3) { allow.tap() }
    }

    func testOnboardingThenEveryTab() {
        let nameField = app.textFields["이름 (예: 홍길동)"]
        XCTAssertTrue(nameField.waitForExistence(timeout: 5), "onboarding sheet should show on fresh install")
        snap("00-onboarding")
        nameField.tap()
        nameField.typeText("Kevin")
        app.buttons["시작하기"].tap()
        allowNotificationsIfAsked()

        let tabs = app.tabBars.firstMatch
        XCTAssertTrue(tabs.waitForExistence(timeout: 5))
        XCTAssertTrue(app.openBank(timeout: 5))
        XCTAssertFalse(app.staticTexts["SIM"].exists, "SIM badge was removed on user request")
        snap("01-home")

        for (tab, expect) in [("상품", "금융 상품"), ("내 방", "미니룸"), ("설정", "내 정보")] {
            tabs.buttons[tab].tap()
            XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS %@", expect)).firstMatch.waitForExistence(timeout: 5),
                          "tab \(tab) should show '\(expect)'")
            snap("tab-\(tab)")
        }
        // the version row at the bottom of 설정 must show the real marketing version + build, not XcodeGen's 1.0 (1)
        let version = app.staticTexts["1.1.0 (2)"]
        for _ in 0..<8 where !version.exists { app.swipeUp() }
        XCTAssertTrue(version.exists, "설정 shows version 1.1.0 (2)")
        snap("settings-version")
    }

    func testChallengeLogAndReport() {
        let nameField = app.textFields["이름 (예: 홍길동)"]
        XCTAssertTrue(nameField.waitForExistence(timeout: 5))
        app.buttons["시작하기"].tap()
        allowNotificationsIfAsked()
        XCTAssertTrue(app.openBank(timeout: 5))
        XCTAssertTrue(app.staticTexts["무지출 챌린지"].exists, "challenge tile on home")
        app.buttons["tile-challenge"].tap()
        let resist = app.buttons["resist-button"]
        XCTAssertTrue(resist.waitForExistence(timeout: 5))
        snap("challenge-empty")
        resist.tap()
        XCTAssertTrue(app.navigationBars["참았다!"].waitForExistence(timeout: 5))
        app.buttons["기록"].tap()
        _ = app.staticTexts["참은 돈 계좌에 입금됐어요"].waitForExistence(timeout: 2)   // success overlay auto-dismisses after 1.2 s
        XCTAssertTrue(app.staticTexts["🔥 연속"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label BEGINSWITH '+'")).firstMatch.exists, "logged entry listed with its amount")
        snap("challenge-logged")
        let trophy = app.staticTexts.matching(NSPredicate(format: "label BEGINSWITH '🏆'")).firstMatch
        if trophy.exists { trophy.tap() }   // achievement banner overlays the back button
        app.navigationBars.buttons.element(boundBy: 0).tap()
        XCTAssertTrue(app.staticTexts["이번 달 소비"].waitForExistence(timeout: 5))
        app.buttons["tile-report"].tap()
        XCTAssertTrue(app.navigationBars["소비 리포트"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts["이 달 지출"].exists)
        snap("report")
    }

    func testScenarioOnboardingPayday() {
        XCTAssertTrue(app.textFields["이름 (예: 홍길동)"].waitForExistence(timeout: 5))
        app.buttons.matching(NSPredicate(format: "label BEGINSWITH '시나리오로 시작'")).firstMatch.tap()
        XCTAssertTrue(app.staticTexts["어떤 삶으로 시작할까요?"].waitForExistence(timeout: 5))
        app.buttons.matching(NSPredicate(format: "label CONTAINS '월급날'")).firstMatch.tap()
        snap("onboarding-scenario")
        app.buttons["scenario-start"].tap()
        allowNotificationsIfAsked()
        XCTAssertTrue(app.openBank(timeout: 8))
        app.openBank(); app.buttons["home-alerts"].tap()
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS '급여'")).firstMatch.waitForExistence(timeout: 5), "payday scenario schedules the salary")
        snap("scenario-events")
    }

    func testQuestCardAnswerAndTransactionHistory() {
        app.terminate()
        app.launchArguments = ["-uitest-reset", "-uitest-quest", "ins_01"]
        app.launch()
        XCTAssertTrue(app.openBank(timeout: 8))
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS '실손의료보험에 가입하시겠습니까'")).firstMatch.waitForExistence(timeout: 5), "quest card on home")
        snap("quest-card")
        app.buttons["quest-option-0"].tap()
        XCTAssertTrue(app.buttons["이렇게 한다"].waitForExistence(timeout: 5))
        app.buttons["이렇게 한다"].tap()
        XCTAssertTrue(app.buttons["확인"].waitForExistence(timeout: 5))
        snap("quest-outcome")
        app.buttons["확인"].tap()
        XCTAssertTrue(app.buttons["life-row"].waitForExistence(timeout: 5))
        app.buttons["life-row"].tap()
        XCTAssertTrue(app.navigationBars["내 인생"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts["재력 단계"].exists)
        snap("life")
        app.navigationBars.buttons.element(boundBy: 0).tap()
        // 전체 거래내역 with period/category filters, then a transaction detail
        let all = app.buttons["전체보기"]
        XCTAssertTrue(all.waitForExistence(timeout: 5))
        all.tap()
        XCTAssertTrue(app.navigationBars["전체 거래내역"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.buttons["12개월"].exists && app.buttons["7일"].exists)
        app.buttons["12개월"].tap()
        snap("history")
        let row = app.buttons.matching(NSPredicate(format: "label CONTAINS '₩'")).firstMatch
        XCTAssertTrue(row.waitForExistence(timeout: 5))
        row.tap()
        XCTAssertTrue(app.navigationBars["거래 상세"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.staticTexts["거래 후 잔액"].exists)
        snap("transaction-detail")
    }

    func testAccountDetailOpens() {
        let nameField = app.textFields["이름 (예: 홍길동)"]
        XCTAssertTrue(nameField.waitForExistence(timeout: 5))
        app.buttons["시작하기"].tap()
        allowNotificationsIfAsked()
        XCTAssertTrue(app.openBank(timeout: 5))
        let firstAccount = app.buttons.matching(NSPredicate(format: "label CONTAINS '입출금통장' AND label CONTAINS '₩'")).firstMatch
        XCTAssertTrue(firstAccount.waitForExistence(timeout: 5), "an account tile with a ₩ balance should exist")
        firstAccount.tap()
        XCTAssertTrue(app.buttons["기간·분류별 보기"].waitForExistence(timeout: 5))
        if isStoreBuild {
            XCTAssertFalse(app.buttons["잔액 설정"].exists, "Store build must not offer manual balance editing")
        } else {
            XCTAssertTrue(app.buttons["잔액 설정"].exists)
        }
        snap("account-detail")
    }
}

extension SmokeTests {
    func testScheduleDepositViaEditor() {
        XCTAssertTrue(app.textFields["이름 (예: 홍길동)"].waitForExistence(timeout: 5))
        app.buttons["시작하기"].tap()
        allowNotificationsIfAsked()
        app.openBank(); app.buttons["home-alerts"].tap()
        XCTAssertTrue(app.staticTexts["예약된 알림이 없습니다"].waitForExistence(timeout: 5))
        if isStoreBuild {   // no hand-made deposits in the Store build: the add button must be gone
            XCTAssertLessThanOrEqual(app.navigationBars["예약 알림"].buttons.count, 1, "Store build must not offer manual scheduled deposits (only the back button)")
            snap("schedule-list-store"); return
        }

        app.navigationBars["예약 알림"].buttons.element(boundBy: app.navigationBars["예약 알림"].buttons.count - 1).tap()
        XCTAssertTrue(app.navigationBars["알림 예약"].waitForExistence(timeout: 5))
        app.buttons.matching(NSPredicate(format: "label BEGINSWITH '엄마 · 용돈'")).firstMatch.tap()
        snap("schedule-editor")
        let save = app.navigationBars["알림 예약"].buttons["저장"]
        XCTAssertTrue(save.isEnabled, "preset should make the form valid")
        save.tap()

        let row = app.staticTexts["엄마 · 용돈"]
        XCTAssertTrue(row.waitForExistence(timeout: 5), "saved event should be listed")
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS '5,000,000'")).firstMatch.exists)
        snap("schedule-list")
    }
}

extension SmokeTests {
    func testTossThemeScreens() {
        app.terminate()
        app.launchArguments = ["-uitest-reset", "-uitest-theme", "toss"]
        app.launch()
        XCTAssertTrue(app.textFields["이름 (예: 홍길동)"].waitForExistence(timeout: 5))
        app.buttons["시작하기"].tap()
        allowNotificationsIfAsked()
        XCTAssertTrue(app.openBank(timeout: 5))
        XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label ENDSWITH '원'")).firstMatch.exists, "TOSS theme writes amounts with 원 suffix")
        snap("toss-home")
        app.buttons["송금"].firstMatch.tap()
        XCTAssertTrue(app.navigationBars["송금"].waitForExistence(timeout: 5))
        snap("toss-transfer")
        app.navigationBars.buttons.element(boundBy: 0).tap()
        app.staticTexts["입출금통장"].firstMatch.tap()
        XCTAssertTrue(app.buttons["보내기"].waitForExistence(timeout: 5), "TOSS account detail shows 채우기/보내기")
        snap("toss-account")
        app.navigationBars.buttons.element(boundBy: 0).tap()
        app.tabBars.firstMatch.buttons["설정"].tap()
        snap("toss-settings")
    }
}

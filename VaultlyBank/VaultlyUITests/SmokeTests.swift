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
        XCTAssertTrue(app.staticTexts["내 계좌"].waitForExistence(timeout: 5))
        XCTAssertTrue(app.otherElements["시뮬레이션 모드"].exists || app.staticTexts["시뮬레이션 모드"].exists || app.staticTexts["SIM"].exists,
                      "SIM badge must be visible on Home")
        snap("01-home")

        for (tab, expect) in [("증권", "내 투자"), ("상품", "금융 상품"), ("예약 알림", "예약 알림"), ("설정", "내 정보")] {
            tabs.buttons[tab].tap()
            XCTAssertTrue(app.staticTexts.matching(NSPredicate(format: "label CONTAINS %@", expect)).firstMatch.waitForExistence(timeout: 5),
                          "tab \(tab) should show '\(expect)'")
            snap("tab-\(tab)")
        }
    }

    func testAccountDetailOpens() {
        let nameField = app.textFields["이름 (예: 홍길동)"]
        XCTAssertTrue(nameField.waitForExistence(timeout: 5))
        app.buttons["시작하기"].tap()
        allowNotificationsIfAsked()
        XCTAssertTrue(app.staticTexts["내 계좌"].waitForExistence(timeout: 5))
        let firstAccount = app.buttons.matching(NSPredicate(format: "label CONTAINS '₩'")).firstMatch
        XCTAssertTrue(firstAccount.waitForExistence(timeout: 5), "an account tile with a ₩ balance should exist")
        firstAccount.tap()
        XCTAssertTrue(app.buttons["잔액 설정"].waitForExistence(timeout: 5))
        snap("account-detail")
    }
}

extension SmokeTests {
    func testScheduleDepositViaEditor() {
        XCTAssertTrue(app.textFields["이름 (예: 홍길동)"].waitForExistence(timeout: 5))
        app.buttons["시작하기"].tap()
        allowNotificationsIfAsked()
        app.tabBars.firstMatch.buttons["예약 알림"].tap()
        XCTAssertTrue(app.staticTexts["예약된 알림이 없습니다"].waitForExistence(timeout: 5))

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
        XCTAssertTrue(app.staticTexts["총자산"].waitForExistence(timeout: 5))
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
        app.tabBars.firstMatch.buttons["증권"].tap()
        XCTAssertTrue(app.staticTexts["기본계좌"].waitForExistence(timeout: 5))
        snap("toss-invest")
        app.tabBars.firstMatch.buttons["설정"].tap()
        snap("toss-settings")
    }
}

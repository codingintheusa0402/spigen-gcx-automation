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

        for (tab, expect) in [("증권", "내 투자"), ("상품", "금융 상품"), ("예약 알림", "예약 알림"), ("설정", "프로필")] {
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

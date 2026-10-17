import XCTest

extension XCUIApplication {
    /// The app opens on the 미니룸 tab; the bank home (계좌·총자산) is the second tab.
    @discardableResult
    func openBank(timeout: TimeInterval = 5) -> Bool {
        let tab = tabBars.firstMatch.buttons["은행"]
        if tab.waitForExistence(timeout: timeout) { tab.tap() }
        return staticTexts["내 계좌"].waitForExistence(timeout: timeout) || staticTexts["총자산"].waitForExistence(timeout: 2)
    }
}

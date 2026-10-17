import SwiftUI
import UIKit

/// Pretendard everywhere — the typeface most Korean fintech apps ship with.
enum AppFont {
    static func name(_ weight: Font.Weight) -> String {
        switch weight {
        case .ultraLight, .thin, .light, .regular: return "Pretendard-Regular"
        case .medium: return "Pretendard-Medium"
        case .semibold: return "Pretendard-SemiBold"
        default: return "Pretendard-Bold"        // .bold, .heavy, .black — real bank apps never go heavier
        }
    }

    static func size(_ style: Font.TextStyle) -> CGFloat {
        switch style {
        case .largeTitle: return 34
        case .title: return 28
        case .title2: return 22
        case .title3: return 20
        case .headline, .body: return 17
        case .callout: return 16
        case .subheadline: return 15
        case .footnote: return 13
        case .caption: return 12
        case .caption2: return 11
        @unknown default: return 17
        }
    }

    static func uiFont(_ weight: Font.Weight, size: CGFloat) -> UIFont {
        UIFont(name: name(weight), size: size) ?? .systemFont(ofSize: size)
    }

    /// Navigation bar titles and tab labels are UIKit-drawn, so they need the font via appearance.
    static func installUIKitAppearance() {
        let nav = UINavigationBar.appearance()
        nav.largeTitleTextAttributes = [.font: uiFont(.bold, size: 34)]
        nav.titleTextAttributes = [.font: uiFont(.semibold, size: 17)]
        UITabBarItem.appearance().setTitleTextAttributes([.font: uiFont(.medium, size: 10)], for: .normal)
        UITabBarItem.appearance().setTitleTextAttributes([.font: uiFont(.semibold, size: 10)], for: .selected)
        UIBarButtonItem.appearance().setTitleTextAttributes([.font: uiFont(.regular, size: 17)], for: .normal)
        UISegmentedControl.appearance().setTitleTextAttributes([.font: uiFont(.medium, size: 13)], for: .normal)
    }
}

extension Font {
    static func app(_ style: Font.TextStyle, _ weight: Font.Weight = .regular) -> Font {
        .custom(AppFont.name(weight), size: AppFont.size(style), relativeTo: style)
    }
    static func app(size: CGFloat, _ weight: Font.Weight = .regular) -> Font {
        .custom(AppFont.name(weight), fixedSize: size)
    }
}

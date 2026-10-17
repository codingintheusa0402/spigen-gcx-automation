import SwiftUI

enum ThemeStyle { case card, toss }

/// Bank brands. `toss` reproduces the flat dark list layout of Korean neobank apps.
struct Theme: Identifiable {
    let id: String
    let bankName: String
    let brokerName: String
    let tagline: String
    let accent: Color
    let accentSecondary: Color
    let background: Color
    let card: Color
    let cornerRadius: CGFloat
    let cardGradient: [Color]
    let cardPalette: [[Color]]
    var style: ThemeStyle = .card
    var wonSuffix = false
    var onAccent: Color = .white   // text drawn on top of `accent`          // 219,320원 instead of ₩219,320
    var displayName: String { style == .toss ? (AppVariant.isStore ? "미니멀 다크" : "TOSS") : bankName }

    static let tossBlue = Color(red: 0.192, green: 0.510, blue: 0.965)   // #3182F6
    static let tossRed = Color(red: 0.941, green: 0.267, blue: 0.322)    // #F04452

    static let all: [Theme] = [
        // Store variant ships fictional names only (App Store 4.1 / 5.2.1); the layout is the same.
        Theme(id: "toss", bankName: AppVariant.isStore ? "미니멀뱅크" : "토스뱅크", brokerName: AppVariant.isStore ? "미니멀증권" : "토스증권", tagline: "금융을 쉽고 간편하게",
              accent: tossBlue, accentSecondary: tossRed,
              background: adaptive(light: (0.949, 0.957, 0.965), dark: (0.067, 0.067, 0.075)),
              card: adaptive(light: (1, 1, 1), dark: (0.110, 0.114, 0.129)),
              cornerRadius: 20,
              cardGradient: [tossBlue, Color(red: 0.106, green: 0.392, blue: 0.855)],
              cardPalette: [[tossBlue, Color(red: 0.106, green: 0.392, blue: 0.855)], [tossRed, Color(red: 0.82, green: 0.23, blue: 0.28)],
                            [Color(red: 0.12, green: 0.72, blue: 0.42), Color(red: 0.08, green: 0.55, blue: 0.33)],
                            [Color(red: 1.0, green: 0.70, blue: 0.0), Color(red: 0.95, green: 0.55, blue: 0.0)],
                            [Color(red: 0.48, green: 0.36, blue: 0.96), Color(red: 0.36, green: 0.25, blue: 0.85)],
                            [Color(red: 0.55, green: 0.58, blue: 0.63), Color(red: 0.35, green: 0.38, blue: 0.43)]],
              style: .toss, wonSuffix: true),
        Theme(id: "nova", bankName: "Nova Bank", brokerName: "Nova 증권", tagline: "심플한 금융의 시작",
              accent: Color(red: 0.0, green: 0.275, blue: 1.0), accentSecondary: Color(red: 0.94, green: 0.27, blue: 0.32),
              background: adaptive(light: (0.961, 0.965, 0.973), dark: (0.071, 0.075, 0.086)),
              card: adaptive(light: (1, 1, 1), dark: (0.110, 0.114, 0.129)),
              cornerRadius: 14,
              cardGradient: [Color(red: 0.0, green: 0.275, blue: 1.0), Color(red: 0.0, green: 0.275, blue: 1.0)],
              cardPalette: flat([Color(red: 0.0, green: 0.275, blue: 1.0), Color(red: 0.0, green: 0.6, blue: 0.78), Color(red: 0.33, green: 0.36, blue: 0.9),
                                 Color(red: 0.15, green: 0.65, blue: 0.4), Color(red: 0.95, green: 0.55, blue: 0.1), Color(red: 0.45, green: 0.48, blue: 0.53)])),
        Theme(id: "heritage", bankName: "Heritage Trust", brokerName: "Heritage 증권", tagline: "Since 1887 · 신뢰의 전통",
              accent: Color(red: 1.0, green: 0.737, blue: 0.0), accentSecondary: Color(red: 0.94, green: 0.27, blue: 0.32),
              background: adaptive(light: (0.953, 0.957, 0.965), dark: (0.075, 0.075, 0.08)),
              card: adaptive(light: (1, 1, 1), dark: (0.118, 0.118, 0.129)),
              cornerRadius: 12,
              cardGradient: [Color(red: 1.0, green: 0.737, blue: 0.0), Color(red: 1.0, green: 0.737, blue: 0.0)],
              cardPalette: flat([Color(red: 0.33, green: 0.33, blue: 0.38), Color(red: 1.0, green: 0.737, blue: 0.0), Color(red: 0.55, green: 0.42, blue: 0.2),
                                 Color(red: 0.2, green: 0.45, blue: 0.4), Color(red: 0.7, green: 0.3, blue: 0.3), Color(red: 0.45, green: 0.48, blue: 0.53)]),
              onAccent: Color(red: 0.15, green: 0.15, blue: 0.17)),
        Theme(id: "orbit", bankName: "Orbit Pay", brokerName: "Orbit 증권", tagline: "지갑 하나로, 어디서나",
              accent: Color(red: 0.0, green: 0.52, blue: 0.52), accentSecondary: Color(red: 0.94, green: 0.27, blue: 0.32),
              background: adaptive(light: (0.957, 0.965, 0.965), dark: (0.067, 0.078, 0.078)),
              card: adaptive(light: (1, 1, 1), dark: (0.11, 0.125, 0.125)),
              cornerRadius: 16,
              cardGradient: [Color(red: 0.0, green: 0.52, blue: 0.52), Color(red: 0.0, green: 0.52, blue: 0.52)],
              cardPalette: flat([Color(red: 0.0, green: 0.52, blue: 0.52), Color(red: 0.0, green: 0.42, blue: 0.7), Color(red: 0.5, green: 0.35, blue: 0.75),
                                 Color(red: 0.2, green: 0.6, blue: 0.35), Color(red: 0.9, green: 0.5, blue: 0.15), Color(red: 0.45, green: 0.48, blue: 0.53)])),
    ]

    static func named(_ id: String) -> Theme { all.first { $0.id == id } ?? all[0] }

    /// Solid colors only — gradients read as fake next to real bank apps.
    static func flat(_ colors: [Color]) -> [[Color]] { colors.map { [$0, $0] } }

    /// Color that follows the current light/dark trait, so every theme works in both modes.
    static func adaptive(light: (Double, Double, Double), dark: (Double, Double, Double)) -> Color {
        Color(UIColor { trait in
            let c = trait.userInterfaceStyle == .dark ? dark : light
            return UIColor(red: c.0, green: c.1, blue: c.2, alpha: 1)
        })
    }
    func gradient(for index: Int) -> LinearGradient {
        let pair = cardPalette[abs(index) % cardPalette.count]
        return LinearGradient(colors: pair, startPoint: .topLeading, endPoint: .bottomTrailing)
    }
}

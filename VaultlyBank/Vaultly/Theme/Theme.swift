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
    var wonSuffix = false          // 219,320원 instead of ₩219,320
    var displayName: String { style == .toss ? "TOSS" : bankName }

    static let tossBlue = Color(red: 0.192, green: 0.510, blue: 0.965)   // #3182F6
    static let tossRed = Color(red: 0.941, green: 0.267, blue: 0.322)    // #F04452

    static let all: [Theme] = [
        Theme(id: "toss", bankName: "토스뱅크", brokerName: "토스증권", tagline: "금융을 쉽고 간편하게",
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
              accent: Color(red: 0.19, green: 0.42, blue: 1.0), accentSecondary: Color(red: 0.0, green: 0.78, blue: 0.9),
              background: Color(UIColor.systemGroupedBackground), card: Color(UIColor.secondarySystemGroupedBackground),
              cornerRadius: 22,
              cardGradient: [Color(red: 0.19, green: 0.42, blue: 1.0), Color(red: 0.36, green: 0.2, blue: 0.95)],
              cardPalette: [[.blue, .indigo], [.cyan, .blue], [.mint, .teal], [.indigo, .purple], [.orange, .pink], [.gray, .black]]),
        Theme(id: "heritage", bankName: "Heritage Trust", brokerName: "Heritage 증권", tagline: "Since 1887 · 신뢰의 전통",
              accent: Color(red: 0.07, green: 0.2, blue: 0.4), accentSecondary: Color(red: 0.8, green: 0.65, blue: 0.3),
              background: adaptive(light: (0.96, 0.95, 0.92), dark: (0.09, 0.09, 0.08)),
              card: adaptive(light: (1, 1, 1), dark: (0.15, 0.15, 0.13)),
              cornerRadius: 12,
              cardGradient: [Color(red: 0.07, green: 0.2, blue: 0.4), Color(red: 0.05, green: 0.1, blue: 0.25)],
              cardPalette: [[Color(red: 0.07, green: 0.2, blue: 0.4), .black], [Color(red: 0.55, green: 0.1, blue: 0.15), .black],
                            [Color(red: 0.1, green: 0.35, blue: 0.3), .black], [Color(red: 0.8, green: 0.65, blue: 0.3), .brown],
                            [.gray, .black], [.brown, .black]]),
        Theme(id: "orbit", bankName: "Orbit Pay", brokerName: "Orbit 증권", tagline: "지갑 하나로, 어디서나",
              accent: Color(red: 0.55, green: 0.25, blue: 0.95), accentSecondary: Color(red: 1.0, green: 0.45, blue: 0.6),
              background: adaptive(light: (0.95, 0.94, 0.98), dark: (0.08, 0.07, 0.12)),
              card: adaptive(light: (1, 1, 1), dark: (0.14, 0.13, 0.2)),
              cornerRadius: 26,
              cardGradient: [Color(red: 0.55, green: 0.25, blue: 0.95), Color(red: 1.0, green: 0.45, blue: 0.6)],
              cardPalette: [[.purple, .pink], [.pink, .orange], [.indigo, .purple], [.teal, .green], [.orange, .yellow], [.gray, .black]]),
    ]

    static func named(_ id: String) -> Theme { all.first { $0.id == id } ?? all[0] }

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

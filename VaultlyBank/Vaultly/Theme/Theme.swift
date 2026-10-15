import SwiftUI

/// Fictional bank brands. None of these copy a real bank's name, logo or layout.
struct Theme: Identifiable {
    let id: String
    let bankName: String
    let tagline: String
    let accent: Color
    let accentSecondary: Color
    let background: Color
    let card: Color
    let fontDesign: Font.Design
    let cornerRadius: CGFloat
    let cardGradient: [Color]
    let cardPalette: [[Color]]

    static let all: [Theme] = [
        Theme(id: "nova", bankName: "Nova Bank", tagline: "심플한 금융의 시작",
              accent: Color(red: 0.19, green: 0.42, blue: 1.0), accentSecondary: Color(red: 0.0, green: 0.78, blue: 0.9),
              background: Color(UIColor.systemGroupedBackground), card: Color(UIColor.secondarySystemGroupedBackground),
              fontDesign: .rounded, cornerRadius: 22,
              cardGradient: [Color(red: 0.19, green: 0.42, blue: 1.0), Color(red: 0.36, green: 0.2, blue: 0.95)],
              cardPalette: [[.blue, .indigo], [.cyan, .blue], [.mint, .teal], [.indigo, .purple], [.orange, .pink], [.gray, .black]]),
        Theme(id: "heritage", bankName: "Heritage Trust", tagline: "Since 1887 · 신뢰의 전통",
              accent: Color(red: 0.07, green: 0.2, blue: 0.4), accentSecondary: Color(red: 0.8, green: 0.65, blue: 0.3),
              background: Color(red: 0.96, green: 0.95, blue: 0.92), card: .white,
              fontDesign: .serif, cornerRadius: 12,
              cardGradient: [Color(red: 0.07, green: 0.2, blue: 0.4), Color(red: 0.05, green: 0.1, blue: 0.25)],
              cardPalette: [[Color(red: 0.07, green: 0.2, blue: 0.4), .black], [Color(red: 0.55, green: 0.1, blue: 0.15), .black],
                            [Color(red: 0.1, green: 0.35, blue: 0.3), .black], [Color(red: 0.8, green: 0.65, blue: 0.3), .brown],
                            [.gray, .black], [.brown, .black]]),
        Theme(id: "orbit", bankName: "Orbit Pay", tagline: "지갑 하나로, 어디서나",
              accent: Color(red: 0.55, green: 0.25, blue: 0.95), accentSecondary: Color(red: 1.0, green: 0.45, blue: 0.6),
              background: Color(red: 0.08, green: 0.07, blue: 0.12), card: Color(red: 0.14, green: 0.13, blue: 0.2),
              fontDesign: .default, cornerRadius: 26,
              cardGradient: [Color(red: 0.55, green: 0.25, blue: 0.95), Color(red: 1.0, green: 0.45, blue: 0.6)],
              cardPalette: [[.purple, .pink], [.pink, .orange], [.indigo, .purple], [.teal, .green], [.orange, .yellow], [.gray, .black]]),
    ]

    static func named(_ id: String) -> Theme { all.first { $0.id == id } ?? all[0] }

    var isDark: Bool { id == "orbit" }
    func gradient(for index: Int) -> LinearGradient {
        let pair = cardPalette[abs(index) % cardPalette.count]
        return LinearGradient(colors: pair, startPoint: .topLeading, endPoint: .bottomTrailing)
    }
}

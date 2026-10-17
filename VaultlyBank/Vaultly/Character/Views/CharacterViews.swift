import SwiftUI
import UIKit

/// Loads the character manifest and slices sprite-sheet frames (cached).
final class CharacterAssets {
    static let shared = CharacterAssets()
    let manifest: CharacterManifest?
    private var sheets: [String: CGImage] = [:]
    private var frames: [String: CGImage] = [:]

    private init() {
        // Vaultly's own art set is the default; the LPC-based set is used only if the original manifest is missing
        manifest = (Bundle.main.url(forResource: "characters_original", withExtension: "json")
                    ?? Bundle.main.url(forResource: "characters", withExtension: "json"))
            .flatMap { try? Data(contentsOf: $0) }
            .flatMap { try? CharacterManifest.load($0) }
    }

    func frame(_ file: String, row: Int, col: Int) -> CGImage? {
        let key = "\(file)#\(row)#\(col)"
        if let f = frames[key] { return f }
        if sheets[file] == nil { sheets[file] = UIImage(named: file)?.cgImage }
        guard let sheet = sheets[file] else { return nil }
        let rect = CGRect(x: col * 64, y: row * 64, width: 64, height: 64)
        guard rect.maxX <= CGFloat(sheet.width), rect.maxY <= CGFloat(sheet.height), let f = sheet.cropping(to: rect) else { return nil }
        if frames.count > 600 { frames.removeAll() }
        frames[key] = f
        return f
    }

    func prop(_ name: String) -> CGImage? { frame("ch_prop_\(name)", row: 0, col: 0) }
}

/// The animated pixel character doing an activity, with its props. Drawn on a 96×96 pixel stage.
struct CharacterSpriteView: View {
    let look: CharacterLook
    let activity: String
    var scale: CGFloat = 2

    var body: some View {
        let a = CharacterAssets.shared
        if let m = a.manifest, let act = m.activities[activity], let info = m.anims[act.anim] {
            let fps: Double = act.anim == "walk" ? 8 : (act.anim == "idle" ? 2 : 1)
            TimelineView(.periodic(from: .now, by: 1 / fps)) { tl in
                let tick = Int(tl.date.timeIntervalSinceReferenceDate * fps)
                let col: Int = {
                    if let f = act.frame { return f }
                    switch act.anim {
                    case "walk": return 1 + tick % max(1, info.frames - 1)     // frame 0 is standing
                    case "idle": return tick % info.frames
                    default: return 0
                    }
                }()
                let row = info.rows == 1 ? 0 : (act.row ?? 2)
                let props = act.props ?? []
                ZStack(alignment: .topLeading) {
                    ForEach(props.filter { $0 == "bed" || $0 == "scooter" }, id: \.self) { p in
                        pixel(a.prop(p)).offset(x: 16 * scale, y: (p == "bed" ? 22 : 20) * scale)
                    }
                    ForEach(look.layers(for: act.anim, in: m), id: \.self) { l in
                        pixel(a.frame(l.file(body: look.body, anim: act.anim, prefix: m.prefix), row: row, col: col)).offset(x: 16 * scale, y: 4 * scale)
                    }
                    ForEach(props.filter { $0 != "bed" && $0 != "scooter" }, id: \.self) { p in
                        pixel(a.prop(p)).offset(x: 16 * scale, y: (p == "desk" || p == "counter" ? 14 : 10) * scale)
                    }
                }
                .frame(width: 96 * scale, height: 96 * scale, alignment: .topLeading)
            }
            .accessibilityElement()
            .accessibilityLabel("캐릭터")
            .accessibilityIdentifier("character-sprite")
        } else {
            Text("🙂").font(.system(size: 48 * scale / 2))
        }
    }

    @ViewBuilder private func pixel(_ cg: CGImage?) -> some View {
        if let cg {
            Image(decorative: cg, scale: 1).interpolation(.none).resizable().frame(width: 64 * scale, height: 64 * scale)
        }
    }
}

extension BankStore {
    /// This week's look for the current tier (outfits rotate weekly), and what the character is doing right now.
    var characterNow: (look: CharacterLook, slot: RoutineSlot)? {
        guard let m = CharacterAssets.shared.manifest else { return nil }
        let now = self.now
        let cal = Calendar.current
        let week = UInt64(truncatingIfNeeded: DailyRoutine.weekIndex(now, calendar: cal))
        let seed = lifestyle.seed ^ (week &* 0x9E37_79B9_7F4A_7C15)
        let comps = cal.dateComponents([.hour, .minute, .weekday], from: now)
        let workday = KoreanHolidays.isWorkday(now, calendar: cal)          // weekends and 공휴일 are days off
        let tier = CharacterTier(rawValue: currentTier.rawValue) ?? .broke
        var slot = DailyRoutine.slot(tier: tier, hour: comps.hour ?? 12, minute: comps.minute ?? 0, workday: workday,
                                     variant: Int(lifestyle.seed % 6))
        if data.life.isJailed(at: now) { slot = RoutineSlot(activity: "jail", label: "구치소 수감 중") }
        else if data.life.isInjured(at: now) { slot = RoutineSlot(activity: "injured", label: "부상 회복 중") }
        return (CharacterLook.make(m, tier: currentTier.rawValue, seed: seed, activity: slot.activity), slot)
    }
}

/// Home card: the character living their day.
struct CharacterCard: View {
    @EnvironmentObject var store: BankStore

    var body: some View {
        if let c = store.characterNow {
            Card {
                HStack(spacing: 12) {
                    CharacterSpriteView(look: c.look, activity: c.slot.activity, scale: 1.25)
                        .background(RoundedRectangle(cornerRadius: 12).fill(Color(red: 0.93, green: 0.91, blue: 0.87)))
                        .clipShape(RoundedRectangle(cornerRadius: 12))
                    VStack(alignment: .leading, spacing: 4) {
                        Text("\(store.currentTier.emoji) \(store.currentTier.title)의 하루").font(.app(.caption, .semibold)).foregroundStyle(.secondary)
                        Text(c.slot.label).font(.app(.headline, .bold)).accessibilityIdentifier("character-activity")
                        Text(store.now, format: .dateTime.hour().minute()).font(.system(.subheadline, design: .monospaced)).foregroundStyle(.secondary)
                        if let h = KoreanHolidays.holiday(on: store.now) {
                            Text("🎌 \(h.name) · 쉬는 날").font(.app(.caption2, .semibold)).foregroundStyle(.red)
                        } else if KoreanHolidays.isWeekend(store.now) {
                            Text("주말").font(.app(.caption2, .semibold)).foregroundStyle(.orange)
                        } else {
                            Text("옷은 매주 바뀌어요").font(.app(.caption2)).foregroundStyle(.tertiary)
                        }
                    }
                    Spacer(minLength: 0)
                    NavigationLink(value: HomeRoute.schedule) {
                        Image(systemName: "calendar").font(.app(.title3)).foregroundStyle(store.theme.accent).padding(8)
                    }
                    .buttonStyle(.plain)
                    .accessibilityLabel("오늘의 일정")
                    .accessibilityIdentifier("character-schedule")
                }
            }
            .accessibilityElement(children: .contain)   // a container: children keep their own identifiers
            .accessibilityIdentifier("character-card")
        }
    }
}

/// Credits for the pixel art, as the LPC licences require (CC-BY-SA 3.0 / OGA-BY 3.0: name the authors, link the work).
struct CharacterCreditsView: View {
    struct Entry: Decodable, Identifiable { let file: String; let authors: [String]; let licenses: [String]; let urls: [String]; var id: String { file } }
    struct Doc: Decodable { let source: String; let used_under: String; let files: [Entry]; let vaultly_original: [String] }
    private let doc: Doc? = Bundle.main.url(forResource: "credits", withExtension: "json")
        .flatMap { try? Data(contentsOf: $0) }.flatMap { try? JSONDecoder().decode(Doc.self, from: $0) }

    var body: some View {
        List {
            Section {
                Text("캐릭터 픽셀 아트는 Liberated Pixel Cup(LPC) 아티스트들의 작품을 사용했습니다. 각 그림은 아래 작가들의 것이며, 원본 라이선스(\(doc?.used_under ?? "CC-BY-SA 3.0 / OGA-BY 3.0"))에 따라 표기합니다. 색을 바꾸거나 고친 그림도 같은 라이선스를 따릅니다.")
                    .font(.app(.footnote))
                if let src = doc?.source, let url = URL(string: String(src.split(separator: " ").last ?? "")) {
                    Link("Universal LPC Spritesheet Character Generator", destination: url).font(.app(.footnote, .semibold))
                }
                Link("CC-BY-SA 3.0 라이선스 전문", destination: URL(string: "https://creativecommons.org/licenses/by-sa/3.0/")!).font(.app(.footnote))
                Link("OGA-BY 3.0 라이선스 전문", destination: URL(string: "https://static.opengameart.org/OGA-BY-3.0.txt")!).font(.app(.footnote))
            }
            ForEach(doc?.files ?? []) { e in
                Section(e.file) {
                    Text(e.authors.joined(separator: ", ")).font(.app(.footnote))
                    Text(e.licenses.joined(separator: " · ")).font(.app(.caption2)).foregroundStyle(.secondary)
                    ForEach(e.urls.prefix(3), id: \.self) { u in
                        if let url = URL(string: u) { Link(u, destination: url).font(.app(.caption2)).lineLimit(1) }
                    }
                }
            }
            Section("Vaultly 오리지널") {
                Text((doc?.vaultly_original ?? []).joined(separator: " · ")).font(.app(.footnote))
            }
        }
        .navigationTitle("픽셀 아트 크레딧")
        .accessibilityIdentifier("character-credits")
    }
}

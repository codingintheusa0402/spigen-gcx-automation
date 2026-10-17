import Foundation

// MARK: - Pixel-art character: manifest, look generator, daily routine (pure Foundation, unit-tested)
//
// Art: Universal LPC Spritesheet Character Generator layers (CC-BY-SA 3.0 / OGA-BY 3.0, credited in-app) recoloured
// by scripts/build_characters.py, plus Vaultly's own pieces (helmet, stethoscope, badge, props).
// Files are flat in the bundle: ch_<item>_<body>_<color>_<anim>.png · ch_custom_<piece>_<body>_<anim>.png · ch_prop_<name>.png

struct CharacterManifest: Decodable {
    struct Anim: Decodable { let frames: Int; let rows: Int; let size: Int }
    struct Layer: Decodable { let z: Int; let slot: String; let bodies: [String: [String: [String]]] }   // body → colour → anims
    struct Activity: Decodable {
        let anim: String
        var row: Int?
        var frame: Int?
        var props: [String]?
        var outfit: [String: OutfitChoice]?
        var helmet: Bool?
        var stethoscope: Bool?
        var badge: Bool?
    }
    /// One [item, colour] or a list of alternatives (first that exists for this body wins).
    enum OutfitChoice: Decodable {
        case one([String]), many([[String]])
        init(from d: Decoder) throws {
            let c = try d.singleValueContainer()
            if let m = try? c.decode([[String]].self) { self = .many(m) } else { self = .one(try c.decode([String].self)) }
        }
        var alternatives: [[String]] { switch self { case .one(let x): return [x]; case .many(let m): return m } }
    }
    struct Tier: Decodable {
        let pools: [String: [String?]]
        let colors: [String: [String]]
        init(from d: Decoder) throws {
            let c = try d.container(keyedBy: AnyKey.self)
            var p: [String: [String?]] = [:], col: [String: [String]] = [:]
            for k in c.allKeys {
                if k.stringValue.hasSuffix("_colors") { col[String(k.stringValue.dropLast(7))] = try c.decode([String].self, forKey: k) }
                else { p[k.stringValue] = try c.decode([String?].self, forKey: k) }
            }
            pools = p; colors = col
        }
    }
    struct Custom: Decodable { let z: Int }
    struct AnyKey: CodingKey {
        var stringValue: String; var intValue: Int? { nil }
        init(stringValue: String) { self.stringValue = stringValue }
        init?(intValue: Int) { nil }
    }

    let anims: [String: Anim]
    let layers: [String: Layer]
    let tiers: [String: Tier]
    let activities: [String: Activity]
    let skins: [String]
    let hairColors: [String]
    let customZ: [String: Int]

    enum CodingKeys: String, CodingKey { case anims, layers, tiers, activities, skins, hairColors, custom }
    init(from d: Decoder) throws {
        let c = try d.container(keyedBy: CodingKeys.self)
        anims = try c.decode([String: Anim].self, forKey: .anims)
        layers = try c.decode([String: Layer].self, forKey: .layers)
        tiers = try c.decode([String: Tier].self, forKey: .tiers)
        activities = try c.decode([String: Activity].self, forKey: .activities)
        skins = try c.decode([String].self, forKey: .skins)
        hairColors = try c.decode([String].self, forKey: .hairColors)
        let cu = try c.nestedContainer(keyedBy: AnyKey.self, forKey: .custom)
        var z: [String: Int] = [:]
        for k in cu.allKeys { if let v = try? cu.decode(Custom.self, forKey: k) { z[k.stringValue] = v.z } }
        customZ = z
    }

    static func load(_ data: Data) throws -> CharacterManifest { try JSONDecoder().decode(CharacterManifest.self, from: data) }
}

/// One drawable layer of a character.
struct CharacterLayer: Equatable, Hashable {
    let item: String      // "body", "tshirt", "custom_helmet" …
    let color: String     // "" for custom pieces
    let z: Int
    func file(body: String, anim: String) -> String {
        item.hasPrefix("custom_") ? "ch_\(item)_\(body)_\(anim)" : "ch_\(item)_\(body)_\(color)_\(anim)"
    }
}

/// What the character looks like (stable for a week, then a new outfit), by tier.
struct CharacterLook: Equatable {
    let body: String                // "male" / "female"
    let layers: [CharacterLayer]    // sorted by z

    /// A fresh random look for `tier` (manifest tier key: broke/working/middle/rich/tycoon), optionally dressed for an activity.
    static func make(_ m: CharacterManifest, tier: String, seed: UInt64, activity: String? = nil) -> CharacterLook {
        var rng = GameRNG(seed: seed)
        guard let pool = m.tiers[tier] else { return CharacterLook(body: "male", layers: []) }
        let heads = (pool.pools["head"] ?? []).compactMap { $0 }.filter { m.layers[$0] != nil }
        let head = heads.isEmpty ? "head_m" : rng.pick(heads)
        let body = head.hasSuffix("_f") ? "female" : "male"
        let skin = rng.pick(m.skins)
        var picked: [(String, String)] = [("body", skin), (head, skin)]

        func available(_ item: String) -> [String: [String]]? { m.layers[item]?.bodies[body] }
        for slot in ["legs", "feet", "torso", "neck", "jacket", "face"] {
            let opts = (pool.pools[slot] ?? [nil]).filter { $0 == nil || available($0!) != nil }
            guard !opts.isEmpty, let item = rng.pick(opts), let cols = available(item) else { continue }
            let wanted = (pool.colors[slot] ?? []).filter { cols[$0] != nil }
            let choices = wanted.isEmpty ? cols.keys.sorted() : wanted
            picked.append((item, rng.pick(choices)))
        }
        let hairs = (pool.pools["hair"] ?? []).compactMap { $0 }.filter { available($0) != nil }
        if !hairs.isEmpty { picked.append((rng.pick(hairs), rng.pick(m.hairColors))) }

        if let act = activity.flatMap({ m.activities[$0] }) {
            for (slot, choice) in act.outfit ?? [:] {
                for alt in choice.alternatives where alt.count == 2 {
                    if available(alt[0])?[alt[1]] != nil {
                        picked.removeAll { m.layers[$0.0]?.slot == slot }
                        picked.append((alt[0], alt[1])); break
                    }
                }
            }
            if act.helmet == true { picked.removeAll { m.layers[$0.0]?.slot == "hair" }; picked.append(("custom_helmet", "")) }
            if act.stethoscope == true { picked.append(("custom_stethoscope", "")) }
            if act.badge == true { picked.append(("custom_badge", "")) }
        }
        var layers: [CharacterLayer] = []
        for (item, color) in picked {
            var z = 200
            if let l = m.layers[item] { z = l.z }
            else if let cz = m.customZ[String(item.dropFirst(7))] { z = cz }      // "custom_" prefix
            layers.append(CharacterLayer(item: item, color: color, z: z))
        }
        layers.sort { a, b in a.z == b.z ? a.item < b.item : a.z < b.z }
        return CharacterLook(body: body, layers: layers)
    }

    /// Layers that have frames for `anim` (e.g. suit jackets have no sitting frames — taken off at the desk).
    func layers(for anim: String, in m: CharacterManifest) -> [CharacterLayer] {
        layers.filter { l in
            if l.item.hasPrefix("custom_") { return true }
            return m.layers[l.item]?.bodies[body]?[l.color]?.contains(anim) ?? false
        }
    }
}

// MARK: - Daily routine (what the character is doing right now)

enum CharacterTier: String { case broke, working, middle, rich, tycoon }

struct RoutineSlot: Equatable { let activity: String; let label: String }

enum DailyRoutine {
    /// Hour-by-hour plan for a tier. `workday` false = weekend/holiday. `variant` picks a job flavour (e.g. doctor vs researcher).
    static func slot(tier: CharacterTier, hour: Int, minute: Int = 0, workday: Bool, variant: Int) -> RoutineSlot {
        let h = Double(hour) + Double(minute) / 60
        func r(_ a: String, _ l: String) -> RoutineSlot { RoutineSlot(activity: a, label: l) }
        if !workday {
            switch tier {
            case .broke:
                if h < 9 { return r("sleep", "늦잠") }
                if h < 13 { return variant % 2 == 0 ? r("delivery", "주말 배달 콜") : r("manual", "주말 일용직") }
                if h < 14 { return r("eat_store", "편의점 삼각김밥") }
                if h < 20 { return r("delivery", "배달 콜 대기") }
                return h < 23 ? r("rest", "고시원에서 휴식") : r("sleep", "취침")
            case .working:
                if h < 10 { return r("sleep", "주말 늦잠") }
                if h < 12 { return r("rest", "빨래·청소") }
                if h < 13 { return r("eat", "김치찌개 백반") }
                return h < 24 ? (h < 23 ? r("rest", "넷플릭스") : r("sleep", "취침")) : r("sleep", "취침")
            case .middle:
                if h < 8 { return r("sleep", "취침") }
                if h < 12 { return r("commute", "아이 학원 픽업") }
                if h < 13 { return r("eat", "가족 외식") }
                return h < 23 ? r("rest", "가족과 휴식") : r("sleep", "취침")
            case .rich, .tycoon:
                if h < 8 { return r("sleep", "취침") }
                if h < 13 { return r("commute", tier == .tycoon ? "골프 라운딩" : "골프 연습장") }
                if h < 15 { return r("eat", "호텔 브런치") }
                return h < 24 ? (h < 23 ? r("rest", "라운지에서 휴식") : r("sleep", "취침")) : r("sleep", "취침")
            }
        }
        switch tier {
        case .broke:      // 일용직·배달, 하루 벌어 하루 산다
            if h < 6 { return r("sleep", "쪽잠") }
            if h < 7 { return r("commute", "인력사무소로") }
            if h < 12 { return variant % 2 == 0 ? r("manual", "공사장 일용직") : r("delivery", "오토바이 배달") }
            if h < 12.75 { return r("eat_store", "편의점 컵라면") }
            if h < 19 { return variant % 2 == 0 ? r("manual", "자재 나르기") : r("delivery", "점심·저녁 피크 배달") }
            if h < 19.5 { return r("eat_store", "편의점 도시락") }
            if h < 23 { return r("rest", "고시원") }
            return r("sleep", "취침")
        case .working:    // 9시 출근, 야근은 기본(22시), 점심은 편의점
            if h < 6.5 { return r("sleep", "취침") }
            if h < 8.75 { return r("commute", "만원 지하철 출근") }
            if h < 12 { return variant % 3 == 0 ? r("store", "매장 근무") : r("office", "사무실 근무") }
            if h < 13 { return r("eat_store", "편의점 점심") }
            if h < 18 { return variant % 3 == 0 ? r("store", "매장 근무") : r("office", "사무실 근무") }
            if h < 22 { return r("office", "야근 (수당 적음)") }
            if h < 23 { return r("commute", "막차 퇴근") }
            return r("sleep", "취침")
        case .middle:     // 의사 또는 대기업 연구원, 18시 정시 퇴근
            if h < 6.5 { return r("sleep", "취침") }
            if h < 8.5 { return r("commute", "자차 출근") }
            if h < 12 { return variant % 2 == 0 ? r("hospital", "외래 진료") : r("lab", "연구소 실험") }
            if h < 13 { return r("eat", "구내식당 한식") }
            if h < 18 { return variant % 2 == 0 ? r("hospital", "병동 회진") : r("lab", "연구 보고서") }
            if h < 19 { return r("commute", "정시 퇴근 18:00") }
            if h < 23 { return r("rest", "가족과 저녁") }
            return r("sleep", "취침")
        case .rich:
            if h < 7.5 { return r("sleep", "취침") }
            if h < 9.5 { return r("commute", "기사님 차로 출근") }
            if h < 12 { return r("office", "임원 회의") }
            if h < 14 { return r("eat", "한정식 점심 미팅") }
            if h < 17 { return r("office", "결재") }
            if h < 23 { return r("rest", "와인 바") }
            return r("sleep", "취침")
        case .tycoon:
            if h < 8 { return r("sleep", "취침") }
            if h < 10 { return r("commute", "전용기 이동") }
            if h < 12 { return r("office", "그룹 전략회의") }
            if h < 14.5 { return r("eat", "호텔 오찬") }
            if h < 16 { return r("rest", "VIP 라운지") }
            if h < 17 { return r("office", "M&A 보고") }
            if h < 24 { return r("rest", "만찬") }
            return r("sleep", "취침")
        }
    }

    /// Week index used to rotate outfits ("looks change periodically").
    static func weekIndex(_ date: Date, calendar: Calendar = .current) -> Int {
        let c = calendar.dateComponents([.yearForWeekOfYear, .weekOfYear], from: date)
        return (c.yearForWeekOfYear ?? 0) * 53 + (c.weekOfYear ?? 0)
    }
}

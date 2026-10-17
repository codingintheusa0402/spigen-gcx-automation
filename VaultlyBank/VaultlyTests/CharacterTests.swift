import XCTest
import Foundation
#if canImport(MiniGames)
@testable import MiniGames
#else
@testable import Vaultly
#endif

/// The pixel-art character: every look the generator can produce must point at sprite files that exist,
/// always be dressed, and every hour of every tier's day must map to an activity the art supports.
final class CharacterTests: XCTestCase {
    static let charDir: URL = URL(fileURLWithPath: #filePath).resolvingSymlinksInPath().deletingLastPathComponent()
        .deletingLastPathComponent().appendingPathComponent("Vaultly/Characters")
    lazy var m: CharacterManifest = try! CharacterManifest.load(Data(contentsOf: Self.charDir.appendingPathComponent("characters.json")))
    let tiers = ["broke", "working", "middle", "rich", "tycoon"]

    func exists(_ name: String) -> Bool { FileManager.default.fileExists(atPath: Self.charDir.appendingPathComponent(name + ".png").path) }

    func testEveryLookIsDrawableAndDressed() {
        var checked = 0
        for tier in tiers {
            for seed in UInt64(1)...150 {
                for act in [nil] + m.activities.keys.sorted() {
                    let look = CharacterLook.make(m, tier: tier, seed: seed, activity: act)
                    XCTAssertTrue(look.layers.contains { $0.item == "body" })
                    XCTAssertTrue(look.layers.contains { m.layers[$0.item]?.slot == "head" })
                    XCTAssertTrue(look.layers.contains { m.layers[$0.item]?.slot == "torso" }, "\(tier) \(seed) \(act ?? "-") has no shirt")
                    XCTAssertTrue(look.layers.contains { m.layers[$0.item]?.slot == "legs" }, "\(tier) \(seed) has no legs")
                    let anim = act.flatMap { m.activities[$0]?.anim } ?? "walk"
                    let drawn = look.layers(for: anim, in: m)
                    XCTAssertTrue(drawn.contains { m.layers[$0.item]?.slot == "torso" }, "shirt must have \(anim) frames (\(tier) \(seed))")
                    for l in drawn { XCTAssertTrue(exists(l.file(body: look.body, anim: anim)), l.file(body: look.body, anim: anim)) }
                    checked += 1
                }
            }
        }
        XCTAssertGreaterThan(checked, 5000)
    }

    func testUniformsAndPropsForWork() {
        let store = CharacterLook.make(m, tier: "working", seed: 3, activity: "store")
        XCTAssertTrue(store.layers.contains { ["vest", "apron"].contains($0.item) }, "store staff wear the vest/apron")
        XCTAssertTrue(store.layers.contains { $0.item == "custom_badge" })
        let rider = CharacterLook.make(m, tier: "broke", seed: 9, activity: "delivery")
        XCTAssertTrue(rider.layers.contains { $0.item == "custom_helmet" })
        XCTAssertFalse(rider.layers.contains { m.layers[$0.item]?.slot == "hair" }, "hair is under the helmet")
        for p in m.activities.values.flatMap({ $0.props ?? [] }) { XCTAssertTrue(exists("ch_prop_\(p)"), p) }
        for piece in ["helmet", "stethoscope", "badge"] {
            for body in ["male", "female"] { for a in m.anims.keys { XCTAssertTrue(exists("ch_custom_\(piece)_\(body)_\(a)")) } }
        }
    }

    func testLooksVaryAndRotateWeekly() {
        let a = CharacterLook.make(m, tier: "middle", seed: 1), b = CharacterLook.make(m, tier: "middle", seed: 2)
        XCTAssertEqual(CharacterLook.make(m, tier: "middle", seed: 1), a, "same seed, same look")
        var distinct = Set<[CharacterLayer]>()
        for s in UInt64(1)...60 { distinct.insert(CharacterLook.make(m, tier: "middle", seed: s).layers) }
        XCTAssertGreaterThan(distinct.count, 50, "many different looks")
        _ = b
        let d1 = Date(timeIntervalSince1970: 1_790_000_000)
        XCTAssertNotEqual(DailyRoutine.weekIndex(d1), DailyRoutine.weekIndex(d1.addingTimeInterval(7 * 86_400)))
        XCTAssertEqual(DailyRoutine.weekIndex(d1), DailyRoutine.weekIndex(d1.addingTimeInterval(3600)))
    }

    func testRoutineCoversEveryHourWithSupportedActivities() {
        for t in [CharacterTier.broke, .working, .middle, .rich, .tycoon] {
            for workday in [true, false] {
                for v in 0..<3 {
                    for h in 0..<24 {
                        for minute in [0, 30] {
                            let s = DailyRoutine.slot(tier: t, hour: h, minute: minute, workday: workday, variant: v)
                            XCTAssertNotNil(m.activities[s.activity], "\(t) \(h):\(minute) → \(s.activity)")
                            XCTAssertFalse(s.label.isEmpty)
                        }
                    }
                }
            }
        }
        // the life described for each tier
        XCTAssertEqual(DailyRoutine.slot(tier: .working, hour: 12, minute: 20, workday: true, variant: 1).activity, "eat_store")
        XCTAssertEqual(DailyRoutine.slot(tier: .working, hour: 21, workday: true, variant: 1).label, "야근 (수당 적음)")
        XCTAssertEqual(DailyRoutine.slot(tier: .middle, hour: 18, minute: 10, workday: true, variant: 0).activity, "commute")
        XCTAssertEqual(DailyRoutine.slot(tier: .middle, hour: 10, workday: true, variant: 0).activity, "hospital")
        XCTAssertEqual(DailyRoutine.slot(tier: .middle, hour: 10, workday: true, variant: 1).activity, "lab")
        XCTAssertEqual(DailyRoutine.slot(tier: .broke, hour: 3, workday: true, variant: 0).activity, "sleep")
    }

    func testCreditsCoverEveryLPCFile() throws {
        let data = try Data(contentsOf: Self.charDir.appendingPathComponent("credits.json"))
        let json = try JSONSerialization.jsonObject(with: data) as! [String: Any]
        let files = json["files"] as! [[String: Any]]
        XCTAssertGreaterThan(files.count, 30)
        for f in files {
            XCTAssertFalse((f["authors"] as? [String] ?? []).isEmpty, "authors for \(f["file"] ?? "")")
            let lic = f["licenses"] as? [String] ?? []
            XCTAssertTrue(lic.contains { $0.hasPrefix("CC-BY") || $0.hasPrefix("OGA-BY") || $0 == "CC0" }, "non-GPL licence for \(f["file"] ?? "")")
        }
    }
}

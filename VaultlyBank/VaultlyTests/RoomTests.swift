import XCTest
@testable import Vaultly

final class RoomTests: XCTestCase {
    func testManifestCoversEveryTierAndEveryImageExists() throws {
        let m = try XCTUnwrap(RoomAssets.shared.manifest)
        XCTAssertEqual(m.canvas, [200, 150])
        for tier in WealthTier.allCases {
            let room = try XCTUnwrap(m.tiers[tier.rawValue], "room for \(tier)")
            XCTAssertNotNil(UIImage(named: room.bg), "background \(room.bg)")
            XCTAssertEqual(room.anchor.count, 2)
            XCTAssertTrue(room.anchor[0] < m.canvas[0] && room.anchor[1] <= m.canvas[1])
            XCTAssertGreaterThanOrEqual(room.props.filter { !$0.isClutter }.count, 8)
            XCTAssertEqual(room.props.filter { $0.isClutter }.count, 2, "bills + bottles when poor")
            for p in room.props {
                XCTAssertNotNil(UIImage(named: "rp_\(p.name)"), "prop image \(p.name)")
                XCTAssertTrue(p.x + p.w <= m.canvas[0] + 2 && p.y + p.h <= m.canvas[1] + 2, "\(p.name) inside the canvas")
            }
        }
    }

    func testProgressAndVisibility() throws {
        XCTAssertEqual(RoomState.progress(netWorth: 2_500_000, tier: .broke), 0.5, accuracy: 0.001)
        XCTAssertEqual(RoomState.progress(netWorth: 25_000_000, tier: .working), 0.5, accuracy: 0.001)
        XCTAssertEqual(RoomState.progress(netWorth: 2_000_000_000, tier: .rich), 0.04, accuracy: 0.001)
        XCTAssertEqual(RoomState.progress(netWorth: 500_000_000_000, tier: .tycoon), 1)
        let m = try XCTUnwrap(RoomAssets.shared.manifest)
        let room = try XCTUnwrap(m.tiers["working"])
        let fresh = RoomState(tier: .working, progress: 0, poor: false).visible(room.props)
        let half = RoomState(tier: .working, progress: 0.5, poor: false).visible(room.props)
        let broke = RoomState(tier: .working, progress: 0, poor: true).visible(room.props)
        XCTAssertTrue(fresh.contains { $0.name == "bed" })
        XCTAssertFalse(fresh.contains { $0.name == "bike" })
        XCTAssertGreaterThan(half.count, fresh.count, "more furniture as net worth grows")
        XCTAssertTrue(broke.contains { $0.name == "bills" } && !fresh.contains { $0.name == "bills" })
    }
}

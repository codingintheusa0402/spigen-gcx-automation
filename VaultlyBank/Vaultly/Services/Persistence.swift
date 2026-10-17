import Foundation

enum Persistence {
    private static var url: URL {
        let dir = FileManager.default.urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
            .appendingPathComponent("Vaultly", isDirectory: true)
        try? FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        return dir.appendingPathComponent("bank.json")
    }

    static func load() -> BankData? {
        guard let raw = try? Data(contentsOf: url) else { return nil }
        let dec = JSONDecoder()
        dec.dateDecodingStrategy = .iso8601
        return try? dec.decode(BankData.self, from: raw)
    }

    static func save(_ data: BankData) {
        let enc = JSONEncoder()
        enc.dateEncodingStrategy = .iso8601
        enc.outputFormatting = [.prettyPrinted, .sortedKeys]
        if let raw = try? enc.encode(data) {
            try? raw.write(to: url, options: .atomic)
        }
    }

    static func reset() { try? FileManager.default.removeItem(at: url) }
}

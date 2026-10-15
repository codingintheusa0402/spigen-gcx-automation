import Foundation

/// Virtual clock. 1x follows real time; higher speeds make simulated time run faster
/// (1,200x = 100 years in one real month). Only moves forward.
struct SimClock: Codable, Equatable {
    var speed: Double = 1
    var anchorReal: Date = Date()
    var anchorVirtual: Date = Date()

    func now(real: Date = Date()) -> Date {
        anchorVirtual.addingTimeInterval(max(0, real.timeIntervalSince(anchorReal)) * speed)
    }

    mutating func setSpeed(_ s: Double, real: Date = Date()) {
        anchorVirtual = now(real: real)
        anchorReal = real
        speed = max(1, s)
    }

    var isRealTime: Bool { speed == 1 && abs(anchorVirtual.timeIntervalSince(anchorReal)) < 120 }
}

enum TimeSpeed: Double, CaseIterable, Identifiable {
    case realtime = 1
    case hourIsDay = 24          // 1시간 = 1일
    case hourIsWeek = 168        // 1시간 = 1주
    case dayIsYear = 365         // 1일 = 1년
    case monthIsCentury = 1200   // 1달 = 100년
    var id: Double { rawValue }
    var title: String {
        switch self {
        case .realtime: return "실시간 (1x)"
        case .hourIsDay: return "1시간 = 1일 (24x)"
        case .hourIsWeek: return "1시간 = 1주 (168x)"
        case .dayIsYear: return "1일 = 1년 (365x)"
        case .monthIsCentury: return "1달 = 100년 (1,200x)"
        }
    }
    var short: String {
        switch self {
        case .realtime: return "1x"
        case .hourIsDay: return "24x"
        case .hourIsWeek: return "168x"
        case .dayIsYear: return "365x"
        case .monthIsCentury: return "1200x"
        }
    }
}

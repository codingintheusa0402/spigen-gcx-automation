import Foundation

/// 대한민국 공휴일 — fixed-date holidays, the lunar ones (설날·추석 3일, 부처님오신날) computed with the Chinese
/// lunisolar calendar, and 대체공휴일 by the current rule (2023~). Pure Foundation, cached per year.
enum KoreanHolidays {
    struct Holiday: Equatable, Identifiable {
        let date: Date          // start of day
        let name: String
        let substitute: Bool    // 대체공휴일
        var id: String { "\(Int(date.timeIntervalSince1970))-\(name)" }
    }

    private static var cache: [String: [Holiday]] = [:]
    private static let lock = NSLock()

    /// Every public holiday of `year`, sorted, substitutes included. Election days / one-off 임시공휴일 are not modelled.
    static func holidays(in year: Int, calendar: Calendar = .current) -> [Holiday] {
        let key = "\(year)-\(calendar.identifier)-\(calendar.timeZone.identifier)"
        lock.lock(); if let c = cache[key] { lock.unlock(); return c }; lock.unlock()
        var cal = calendar
        cal.locale = Locale(identifier: "ko_KR")

        func day(_ m: Int, _ d: Int) -> Date? { cal.date(from: DateComponents(year: year, month: m, day: d)).map(cal.startOfDay(for:)) }
        var list: [(Date, String, Bool)] = []   // date, name, eligible for 대체공휴일
        for (m, d, name, sub) in [(1, 1, "신정", false), (3, 1, "삼일절", true), (5, 5, "어린이날", true), (6, 6, "현충일", false),
                                  (8, 15, "광복절", true), (10, 3, "개천절", true), (10, 9, "한글날", true), (12, 25, "성탄절", true)] {
            if let dt = day(m, d) { list.append((dt, name, sub)) }
        }
        // Lunar holidays: walk the Gregorian year once and read the Chinese calendar components of each day.
        var chinese = Calendar(identifier: .chinese)
        chinese.timeZone = cal.timeZone
        if let jan1 = day(1, 1), let dec31 = day(12, 31) {
            var d = jan1
            while d <= dec31 {
                let c = chinese.dateComponents([.month, .day, .isLeapMonth], from: d)
                if c.isLeapMonth != true {
                    switch (c.month, c.day) {
                    case (1, 1):
                        for (off, n) in [(-1, "설날 연휴"), (0, "설날"), (1, "설날 연휴")] {
                            if let x = cal.date(byAdding: .day, value: off, to: d), cal.component(.year, from: x) == year { list.append((x, n, true)) }
                        }
                    case (8, 15):
                        for (off, n) in [(-1, "추석 연휴"), (0, "추석"), (1, "추석 연휴")] {
                            if let x = cal.date(byAdding: .day, value: off, to: d), cal.component(.year, from: x) == year { list.append((x, n, true)) }
                        }
                    case (4, 8): list.append((d, "부처님오신날", true))
                    default: break
                    }
                }
                guard let n = cal.date(byAdding: .day, value: 1, to: d) else { break }
                d = n
            }
        }
        list.sort { $0.0 < $1.0 }

        // 대체공휴일: a holiday on a weekend (설·추석: on Sunday) or on another holiday moves to the next free weekday.
        var taken = Set(list.map { $0.0 })
        var out: [Holiday] = list.map { Holiday(date: $0.0, name: $0.1, substitute: false) }
        var seen = Set<Date>()
        for (dt, name, eligible) in list {
            let dup = seen.contains(dt); seen.insert(dt)
            guard eligible else { continue }
            let wd = cal.component(.weekday, from: dt)
            let isSeollalChuseok = name.hasPrefix("설날") || name.hasPrefix("추석")
            let clash = dup || (isSeollalChuseok ? wd == 1 : (wd == 1 || wd == 7))
            guard clash else { continue }
            // after the last day of a 3-day 연휴, otherwise after the holiday itself
            var cursor = dt
            if isSeollalChuseok { cursor = list.filter { $0.1.hasPrefix(String(name.prefix(2))) }.map { $0.0 }.max() ?? dt }
            while true {
                guard let next = cal.date(byAdding: .day, value: 1, to: cursor) else { break }
                cursor = next
                let w = cal.component(.weekday, from: cursor)
                if w != 1 && w != 7 && !taken.contains(cursor) {
                    taken.insert(cursor)
                    out.append(Holiday(date: cursor, name: "대체공휴일 (\(name.replacingOccurrences(of: " 연휴", with: "")))", substitute: true))
                    break
                }
            }
        }
        out.sort { $0.date < $1.date }
        lock.lock(); cache[key] = out; lock.unlock()
        return out
    }

    static func holiday(on date: Date, calendar: Calendar = .current) -> Holiday? {
        let d = calendar.startOfDay(for: date)
        return holidays(in: calendar.component(.year, from: d), calendar: calendar).first { $0.date == d }
    }
    static func isHoliday(_ date: Date, calendar: Calendar = .current) -> Bool { holiday(on: date, calendar: calendar) != nil }
    static func isWeekend(_ date: Date, calendar: Calendar = .current) -> Bool {
        let w = calendar.component(.weekday, from: date); return w == 1 || w == 7
    }
    /// Mon–Fri and not a public holiday.
    static func isWorkday(_ date: Date, calendar: Calendar = .current) -> Bool { !isWeekend(date, calendar: calendar) && !isHoliday(date, calendar: calendar) }

    /// The next holiday strictly after `date` (looks ahead up to 2 years).
    static func next(after date: Date, calendar: Calendar = .current) -> Holiday? {
        let d = calendar.startOfDay(for: date)
        let y = calendar.component(.year, from: d)
        return (holidays(in: y, calendar: calendar) + holidays(in: y + 1, calendar: calendar)).first { $0.date > d }
    }
}

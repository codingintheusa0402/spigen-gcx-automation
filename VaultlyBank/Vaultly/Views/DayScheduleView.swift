import SwiftUI

/// 오늘의 일정: the character's hour-by-hour day for the current tier, public holidays (공휴일) and what is
/// coming up next. Jail and injury override the routine.
struct DayScheduleView: View {
    @EnvironmentObject var store: BankStore
    @State private var dayOffset = 0

    private var day: Date { Calendar.current.date(byAdding: .day, value: dayOffset, to: store.now) ?? store.now }

    var body: some View {
        ScrollView {
            VStack(spacing: 14) {
                header
                scheduleCard
                holidaysCard
            }
            .padding(16)
        }
        .background(store.theme.background.ignoresSafeArea())
        .navigationTitle("오늘의 일정")
        .navigationBarTitleDisplayMode(.inline)
        .toolbar {
            ToolbarItem(placement: .topBarTrailing) {
                HStack {
                    Button { dayOffset -= 1 } label: { Image(systemName: "chevron.left") }.accessibilityIdentifier("schedule-prev")
                    Button { dayOffset += 1 } label: { Image(systemName: "chevron.right") }.accessibilityIdentifier("schedule-next")
                }
            }
        }
    }

    private var header: some View {
        let holiday = KoreanHolidays.holiday(on: day)
        let weekend = KoreanHolidays.isWeekend(day)
        return Card {
            VStack(alignment: .leading, spacing: 6) {
                HStack {
                    Text(day.formatted(.dateTime.month(.wide).day().weekday(.wide).locale(Locale(identifier: "ko_KR")))).font(.app(.title3, .bold))
                    Spacer()
                    if dayOffset == 0 { Text("오늘").font(.app(.caption, .semibold)).foregroundStyle(store.theme.accent) }
                }
                if let h = holiday {
                    Label("\(h.name) · 공휴일", systemImage: "flag.fill").font(.app(.subheadline, .semibold)).foregroundStyle(.red)
                        .accessibilityIdentifier("schedule-holiday")
                } else if weekend {
                    Label("주말", systemImage: "sun.max.fill").font(.app(.subheadline, .semibold)).foregroundStyle(.orange)
                } else {
                    Label("평일 · \(store.currentTier.title)의 하루", systemImage: "briefcase.fill").font(.app(.subheadline, .semibold)).foregroundStyle(.secondary)
                }
                if store.isJailed { Label("구치소 수감 중 · \(store.life.jailedUntil?.formatted(date: .abbreviated, time: .omitted) ?? "")까지", systemImage: "lock.fill").font(.app(.caption)).foregroundStyle(.red) }
                else if store.isInjured { Label("부상 회복 중 · \(store.life.injuredUntil?.formatted(date: .abbreviated, time: .omitted) ?? "")까지", systemImage: "bandage.fill").font(.app(.caption)).foregroundStyle(.orange) }
            }
        }
    }

    private var scheduleCard: some View {
        Card {
            VStack(alignment: .leading, spacing: 0) {
                Text("시간표").font(.app(.headline, .semibold)).padding(.bottom, 8)
                ForEach(store.schedule(for: day), id: \.start) { s in
                    HStack(alignment: .top, spacing: 12) {
                        Text(s.start).font(.system(.caption, design: .monospaced).weight(.semibold)).frame(width: 44, alignment: .leading)
                            .foregroundStyle(s.isNow ? store.theme.accent : .secondary)
                        Circle().fill(s.isNow ? store.theme.accent : Color.secondary.opacity(0.3)).frame(width: 8, height: 8).padding(.top, 5)
                        VStack(alignment: .leading, spacing: 2) {
                            Text(s.label).font(.app(.subheadline, s.isNow ? .bold : .regular))
                            Text(s.activityName).font(.app(.caption2)).foregroundStyle(.tertiary)
                        }
                        Spacer()
                        if s.isNow { Text("지금").font(.app(.caption2, .semibold)).foregroundStyle(store.theme.accent) }
                    }
                    .padding(.vertical, 6)
                    .accessibilityElement(children: .combine)
                }
            }
        }
        .accessibilityIdentifier("schedule-list")
    }

    private var holidaysCard: some View {
        let year = Calendar.current.component(.year, from: day)
        let list = KoreanHolidays.holidays(in: year)
        let next = KoreanHolidays.next(after: day)
        return Card {
            VStack(alignment: .leading, spacing: 8) {
                HStack {
                    Text("\(String(year))년 공휴일").font(.app(.headline, .semibold))
                    Spacer()
                    if let n = next, let days = Calendar.current.dateComponents([.day], from: Calendar.current.startOfDay(for: day), to: n.date).day {
                        Text("다음 \(n.name) D-\(days)").font(.app(.caption, .semibold)).foregroundStyle(store.theme.accent).accessibilityIdentifier("schedule-next-holiday")
                    }
                }
                ForEach(list) { h in
                    HStack {
                        Text(h.date.formatted(.dateTime.month(.twoDigits).day(.twoDigits).weekday(.short).locale(Locale(identifier: "ko_KR"))))
                            .font(.system(.caption, design: .monospaced)).foregroundStyle(.secondary)
                        Text(h.name).font(.app(.subheadline)).foregroundStyle(h.date < Calendar.current.startOfDay(for: day) ? .tertiary : .primary)
                        Spacer()
                        if h.substitute { Text("대체").font(.app(.caption2, .semibold)).foregroundStyle(.orange) }
                    }
                }
            }
        }
    }
}

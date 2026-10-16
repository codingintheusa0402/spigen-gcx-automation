import SwiftUI
import Charts

/// 소비 리포트: this month's spending by category, top merchants, daily bars, vs last month.
struct ReportView: View {
    @EnvironmentObject var store: BankStore
    @State private var month: Date = Date()

    private var report: SpendingReport { store.report(for: month) }
    private var palette: [Color] {
        [store.theme.accent, .orange, .green, .purple, .pink, .teal, .indigo, .brown, .cyan, .mint, .yellow, .red, .gray, .blue, .secondary, .primary, .secondary]
    }

    var body: some View {
        ScrollView {
            VStack(spacing: 16) {
                monthPicker
                summaryCard
                if !report.slices.isEmpty { categoryCard; merchantsCard; dailyCard }
                else { Card { Text("이 달에는 지출 내역이 없어요.").foregroundStyle(.secondary) } }
            }
            .padding()
        }
        .background(store.theme.background)
        .navigationTitle("소비 리포트")
        .navigationBarTitleDisplayMode(.inline)
        .onAppear { month = store.now }
    }

    private var monthPicker: some View {
        HStack {
            Button { month = Calendar.current.date(byAdding: .month, value: -1, to: month) ?? month } label: { Image(systemName: "chevron.left") }
            Spacer()
            Text(month.formatted(.dateTime.year().month(.wide))).font(.app(.headline, .semibold))
            Spacer()
            Button { month = Calendar.current.date(byAdding: .month, value: 1, to: month) ?? month } label: { Image(systemName: "chevron.right") }
                .disabled(Calendar.current.isDate(month, equalTo: store.now, toGranularity: .month))
        }
        .padding(.horizontal, 4)
    }

    private var summaryCard: some View {
        Card {
            VStack(alignment: .leading, spacing: 8) {
                Text("이 달 지출").font(.app(.subheadline)).foregroundStyle(.secondary)
                MoneyText(amount: report.spent, font: .app(size: 32, .semibold))
                HStack(spacing: 6) {
                    if let r = report.changeRatio {
                        Image(systemName: r >= 0 ? "arrow.up.right" : "arrow.down.right")
                        Text(String(format: "지난달보다 %@%.0f%%", r >= 0 ? "+" : "", r * 100))
                    } else { Text("지난달 비교 데이터 없음") }
                }
                .font(.app(.caption)).foregroundStyle(report.changeRatio ?? 0 > 0 ? store.theme.accentSecondary : .secondary)
                Divider()
                HStack {
                    VStack(alignment: .leading, spacing: 2) {
                        Text("수입").font(.app(.caption)).foregroundStyle(.secondary)
                        MoneyText(amount: report.income, font: .app(.subheadline, .semibold))
                    }
                    Spacer()
                    VStack(alignment: .trailing, spacing: 2) {
                        Text("남은 돈").font(.app(.caption)).foregroundStyle(.secondary)
                        MoneyText(amount: report.income - report.spent, signed: true, font: .app(.subheadline, .semibold))
                            .foregroundStyle(report.income >= report.spent ? store.theme.accent : store.theme.accentSecondary)
                    }
                }
            }
        }
    }

    private var categoryCard: some View {
        Card {
            VStack(alignment: .leading, spacing: 12) {
                Text("카테고리별").font(.app(.headline, .semibold))
                Chart(Array(report.slices.prefix(8).enumerated()), id: \.element.id) { i, s in
                    SectorMark(angle: .value("금액", s.amount.doubleValue), innerRadius: .ratio(0.62), angularInset: 1.5)
                        .foregroundStyle(palette[i % palette.count])
                        .cornerRadius(3)
                }
                .frame(height: 180)
                .overlay {
                    VStack(spacing: 2) {
                        Text(report.slices.first?.category.title ?? "").font(.app(.caption)).foregroundStyle(.secondary)
                        if let top = report.slices.first, report.spent > 0 {
                            Text("\(Int((top.amount / report.spent).doubleValue * 100))%").font(.app(.title3, .bold))
                        }
                    }
                }
                ForEach(Array(report.slices.enumerated()), id: \.element.id) { i, s in
                    HStack(spacing: 10) {
                        Image(systemName: s.category.symbol).foregroundStyle(palette[i % palette.count]).frame(width: 22)
                        Text(s.category.title).font(.app(.subheadline))
                        Text("\(s.count)건").font(.app(.caption)).foregroundStyle(.tertiary)
                        Spacer()
                        MoneyText(amount: s.amount, font: .app(.subheadline, .semibold))
                    }
                    GeometryReader { geo in
                        Capsule().fill(Color.primary.opacity(0.06))
                            .overlay(alignment: .leading) {
                                Capsule().fill(palette[i % palette.count])
                                    .frame(width: geo.size.width * CGFloat(report.spent > 0 ? (s.amount / report.spent).doubleValue : 0))
                            }
                    }
                    .frame(height: 5)
                }
            }
        }
    }

    private var merchantsCard: some View {
        Card {
            VStack(alignment: .leading, spacing: 10) {
                Text("많이 쓴 곳").font(.app(.headline, .semibold))
                ForEach(Array(report.merchants.prefix(5).enumerated()), id: \.element.id) { i, m in
                    HStack {
                        Text("\(i + 1)").font(.app(.caption, .bold)).foregroundStyle(.secondary).frame(width: 18)
                        VStack(alignment: .leading, spacing: 2) {
                            Text(m.name).font(.app(.subheadline, .semibold)).lineLimit(1)
                            Text("\(m.count)건").font(.app(.caption)).foregroundStyle(.secondary)
                        }
                        Spacer()
                        MoneyText(amount: m.amount, font: .app(.subheadline, .semibold))
                    }
                }
            }
        }
    }

    private var dailyCard: some View {
        Card {
            VStack(alignment: .leading, spacing: 10) {
                Text("일별 지출").font(.app(.headline, .semibold))
                Chart(report.dailySpent.sorted { $0.key < $1.key }, id: \.key) { day, amount in
                    BarMark(x: .value("날짜", day, unit: .day), y: .value("지출", amount.doubleValue))
                        .foregroundStyle(store.theme.accent)
                        .cornerRadius(2)
                }
                .chartXAxis { AxisMarks(values: .stride(by: .day, count: 7)) { _ in AxisGridLine(); AxisValueLabel(format: .dateTime.day()) } }
                .chartYAxis { AxisMarks(position: .leading) { v in AxisGridLine(); AxisValueLabel { if let d = v.as(Double.self) { Text(short(d)) } } } }
                .frame(height: 160)
            }
        }
    }

    private func short(_ v: Double) -> String {
        if v >= 100_000_000 { return String(format: "%.0f억", v / 100_000_000) }
        if v >= 10_000 { return String(format: "%.0f만", v / 10_000) }
        return String(format: "%.0f", v)
    }
}

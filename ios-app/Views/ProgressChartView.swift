import SwiftUI
import Charts

struct ProgressChartView: View {
    let done: Int
    let pending: Int

    var body: some View {
        Chart {
            SectorMark(
                angle: .value("Concluídas", done),
                innerRadius: .ratio(0.55),
                angularInset: 2
            )
            .foregroundStyle(.green)

            SectorMark(
                angle: .value("Pendentes", pending),
                innerRadius: .ratio(0.55),
                angularInset: 2
            )
            .foregroundStyle(.orange)
        }
        .frame(height: 220)
    }
}

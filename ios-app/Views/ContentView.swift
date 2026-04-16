import SwiftUI

struct ContentView: View {
    @StateObject var viewModel: DashboardViewModel

    // Em produção, substitua por token obtido via OAuth.
    @State private var accessToken = "COLOQUE_SEU_ACCESS_TOKEN"

    var body: some View {
        NavigationStack {
            VStack(spacing: 16) {
                HStack {
                    Text("Progresso do dia")
                        .font(.title2.bold())
                    Spacer()
                    Text("\(Int(viewModel.completionRate * 100))%")
                        .font(.headline)
                }

                ProgressChartView(
                    done: viewModel.tasks.filter(\.isDone).count,
                    pending: viewModel.tasks.filter { !$0.isDone }.count
                )

                List(viewModel.tasks) { task in
                    Button {
                        viewModel.toggleTask(task)
                    } label: {
                        HStack {
                            Image(systemName: task.isDone ? "checkmark.circle.fill" : "circle")
                                .foregroundStyle(task.isDone ? .green : .gray)

                            VStack(alignment: .leading) {
                                Text(task.title)
                                Text(task.scheduledDate, style: .time)
                                    .font(.caption)
                                    .foregroundStyle(.secondary)
                            }
                        }
                    }
                    .buttonStyle(.plain)
                }
                .listStyle(.plain)

                if let error = viewModel.errorMessage {
                    Text(error)
                        .font(.footnote)
                        .foregroundStyle(.red)
                }
            }
            .padding()
            .navigationTitle("Agenda Checklist")
            .toolbar {
                ToolbarItem(placement: .topBarTrailing) {
                    Button("Sincronizar") {
                        Task {
                            await viewModel.loadToday(accessToken: accessToken)
                        }
                    }
                    .disabled(viewModel.isLoading)
                }
            }
            .task {
                await viewModel.loadToday(accessToken: accessToken)
            }
        }
    }
}

import Foundation

@MainActor
final class DashboardViewModel: ObservableObject {
    @Published var tasks: [DailyTask] = []
    @Published var isLoading = false
    @Published var errorMessage: String?

    private let service: CalendarServiceProtocol

    init(service: CalendarServiceProtocol) {
        self.service = service
    }

    var completionRate: Double {
        guard !tasks.isEmpty else { return 0 }
        let doneCount = tasks.filter(\.isDone).count
        return Double(doneCount) / Double(tasks.count)
    }

    func loadToday(accessToken: String) async {
        isLoading = true
        errorMessage = nil

        defer { isLoading = false }

        do {
            let events = try await service.fetchEvents(for: Date(), accessToken: accessToken)
            tasks = events.map {
                DailyTask(id: $0.id, title: $0.title, scheduledDate: $0.startDate, isDone: false)
            }
        } catch {
            errorMessage = "Não foi possível carregar os compromissos: \(error.localizedDescription)"
        }
    }

    func toggleTask(_ task: DailyTask) {
        guard let idx = tasks.firstIndex(where: { $0.id == task.id }) else { return }
        tasks[idx].isDone.toggle()
    }
}

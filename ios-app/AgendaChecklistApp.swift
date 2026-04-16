import SwiftUI

@main
struct AgendaChecklistApp: App {
    var body: some Scene {
        WindowGroup {
            ContentView(viewModel: DashboardViewModel(service: GoogleCalendarService()))
        }
    }
}

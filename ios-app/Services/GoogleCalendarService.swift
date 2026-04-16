import Foundation

protocol CalendarServiceProtocol {
    func fetchEvents(for day: Date, accessToken: String) async throws -> [AgendaEvent]
}

final class GoogleCalendarService: CalendarServiceProtocol {
    private let session: URLSession

    init(session: URLSession = .shared) {
        self.session = session
    }

    func fetchEvents(for day: Date, accessToken: String) async throws -> [AgendaEvent] {
        var calendar = Calendar.current
        calendar.timeZone = TimeZone.current

        let dayStart = calendar.startOfDay(for: day)
        guard let dayEnd = calendar.date(byAdding: .day, value: 1, to: dayStart) else {
            return []
        }

        let timeMin = ISO8601DateFormatter().string(from: dayStart)
        let timeMax = ISO8601DateFormatter().string(from: dayEnd)

        var components = URLComponents(string: "https://www.googleapis.com/calendar/v3/calendars/primary/events")
        components?.queryItems = [
            URLQueryItem(name: "singleEvents", value: "true"),
            URLQueryItem(name: "orderBy", value: "startTime"),
            URLQueryItem(name: "timeMin", value: timeMin),
            URLQueryItem(name: "timeMax", value: timeMax)
        ]

        guard let url = components?.url else {
            throw URLError(.badURL)
        }

        var request = URLRequest(url: url)
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")

        let (data, response) = try await session.data(for: request)

        guard let httpResponse = response as? HTTPURLResponse, (200...299).contains(httpResponse.statusCode) else {
            throw URLError(.badServerResponse)
        }

        let decoded = try JSONDecoder().decode(CalendarEventsResponse.self, from: data)
        return decoded.items
    }
}

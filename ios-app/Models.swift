import Foundation

struct AgendaEvent: Identifiable, Decodable {
    let id: String
    let title: String
    let startDate: Date
    let endDate: Date?

    enum CodingKeys: String, CodingKey {
        case id
        case summary
        case start
        case end
    }

    enum DateCodingKeys: String, CodingKey {
        case dateTime
        case date
    }

    init(from decoder: Decoder) throws {
        let container = try decoder.container(keyedBy: CodingKeys.self)
        id = try container.decode(String.self, forKey: .id)
        title = try container.decodeIfPresent(String.self, forKey: .summary) ?? "Sem título"

        let startContainer = try container.nestedContainer(keyedBy: DateCodingKeys.self, forKey: .start)
        let endContainer = try? container.nestedContainer(keyedBy: DateCodingKeys.self, forKey: .end)

        startDate = try AgendaEvent.decodeGoogleDate(from: startContainer)

        if let endContainer {
            endDate = try? AgendaEvent.decodeGoogleDate(from: endContainer)
        } else {
            endDate = nil
        }
    }

    private static func decodeGoogleDate(from container: KeyedDecodingContainer<DateCodingKeys>) throws -> Date {
        let formatter = ISO8601DateFormatter()

        if let dateTime = try? container.decode(String.self, forKey: .dateTime),
           let parsed = formatter.date(from: dateTime) {
            return parsed
        }

        if let dateOnly = try? container.decode(String.self, forKey: .date),
           let parsed = ISO8601DateFormatter().date(from: dateOnly + "T00:00:00Z") {
            return parsed
        }

        throw DecodingError.dataCorrupted(.init(codingPath: container.codingPath, debugDescription: "Formato de data inválido do Google Calendar"))
    }
}

struct DailyTask: Identifiable {
    let id: String
    let title: String
    let scheduledDate: Date
    var isDone: Bool
}

struct CalendarEventsResponse: Decodable {
    let items: [AgendaEvent]
}

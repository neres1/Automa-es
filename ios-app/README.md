# App iOS: Agenda → Checklist Diário

Este diretório contém um esqueleto de app em **SwiftUI** para iOS que:

1. Conecta com o **Google Calendar**.
2. Converte compromissos do dia em um **checklist de tarefas**.
3. Exibe um **gráfico de progresso diário** com base no que foi concluído.

## Arquitetura (MVP)

- `AgendaChecklistApp.swift`: ponto de entrada.
- `Models.swift`: modelos de domínio (`AgendaEvent`, `DailyTask`).
- `Services/GoogleCalendarService.swift`: integração com Google Calendar (REST API).
- `ViewModels/DashboardViewModel.swift`: regra de negócio de checklist e progresso.
- `Views/ContentView.swift`: tela principal.
- `Views/ProgressChartView.swift`: gráfico de progresso usando `Charts`.

## Como integrar com Google Calendar

1. No [Google Cloud Console](https://console.cloud.google.com/):
   - Crie um projeto.
   - Ative a **Google Calendar API**.
   - Configure OAuth para iOS.
2. Instale as dependências que preferir para OAuth (`AppAuth` / `GoogleSignIn`), se necessário.
3. Obtenha o **access token** OAuth e injete no `GoogleCalendarService`.
4. Garanta o escopo mínimo:
   - `https://www.googleapis.com/auth/calendar.readonly`

## Regra de checklist

A conversão atual é direta: cada compromisso do dia vira uma tarefa com:
- título = título do evento;
- horário previsto = horário de início do evento;
- concluída = `false` inicialmente.

Você pode evoluir com regras como:
- gerar subtarefas por categoria do evento;
- prioridade por horário de início;
- sugestões automáticas com IA.

## Próximos passos sugeridos

- Persistência local (SwiftData/CoreData) para manter tarefas marcadas offline.
- Sincronização bidirecional opcional (checklist → notas do evento).
- Notificações locais para lembrar tarefas não concluídas.
- Filtros por calendário (trabalho, pessoal, estudos).

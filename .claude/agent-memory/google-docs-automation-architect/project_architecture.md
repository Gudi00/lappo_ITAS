---
name: project_architecture
description: Initial architecture decisions — GUI framework choice (tkinter), module layout, placeholder system ${...}, offline-first design
type: project
---

Принятые архитектурные решения (2026-03-24):

- **GUI**: tkinter (встроенный в Python, не требует установки)
- **Модули**: config, database, google_sheets, google_docs, template_engine, email_sender, gui, main
- **Плейсхолдеры**: формат `${column_name}` для значений и `${table}` для полной таблицы
- **Offline-first**: данные кэшируются в SQLite, шаблоны скачиваются локально, генерация документов работает без интернета
- **Авторизация Google**: OAuth2 (Desktop app flow), токен сохраняется в credentials/token.json
- **Email**: Gmail SMTP с App Password, поддержка тестового режима (dry run)
- **Все технологии бесплатные**: никаких платных сервисов

**Why:** Преподаватель не должен платить за инструменты и не должен зависеть от постоянного интернета.
**How to apply:** Любые новые функции должны работать offline после первичной синхронизации данных.

# Инструкция по настройке

## 1. Установка Python

Убедитесь, что установлен Python 3.10 или новее:
```bash
python --version
```

## 2. Создание виртуального окружения

```bash
cd /home/gudi/Documents/pythonProjectsLinux/lappo_ideal
python -m venv venv
source venv/bin/activate   # Linux/Mac
# venv\Scripts\activate    # Windows
```

## 3. Установка зависимостей

```bash
pip install -r requirements.txt
```

## 4. Настройка Google Cloud Console

### 4.1. Создание проекта
1. Перейдите на https://console.cloud.google.com/
2. Нажмите "Создать проект"
3. Дайте проекту любое имя (например, "Автоматизация документов")

### 4.2. Включение API
1. В меню слева выберите "API и сервисы" -> "Библиотека"
2. Найдите и включите:
   - **Google Sheets API**
   - **Google Drive API**

### 4.3. Создание учётных данных OAuth 2.0
1. "API и сервисы" -> "Учётные данные"
2. "Создать учётные данные" -> "Идентификатор клиента OAuth"
3. Если попросят настроить экран согласия:
   - Тип: "Внешний"
   - Заполните обязательные поля (название, email)
   - Добавьте области доступа:
     - `https://www.googleapis.com/auth/spreadsheets.readonly`
     - `https://www.googleapis.com/auth/drive.readonly`
   - Добавьте свой email в "Тестовые пользователи"
4. Тип приложения: "Приложение для ПК" (Desktop app)
5. Скачайте JSON файл
6. Переименуйте в `credentials.json`
7. Поместите в папку `credentials/`

## 5. Настройка Gmail (для отправки писем)

### 5.1. Включение двухфакторной аутентификации
1. https://myaccount.google.com/security
2. Включите "Двухэтапная аутентификация"

### 5.2. Создание пароля приложения
1. https://myaccount.google.com/apppasswords
2. Выберите "Почта" и "Компьютер Windows"
3. Нажмите "Создать"
4. Скопируйте 16-символьный пароль

### 5.3. Настройка .env файла
```bash
cp .env.example .env
```

Отредактируйте `.env`:
```
GMAIL_ADDRESS=ваш_адрес@gmail.com
GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
```

## 6. Запуск

```bash
python main.py
```

При первом запуске откроется окно браузера для авторизации в Google.
Разрешите доступ — токен сохранится локально и повторная авторизация
не потребуется.

## Структура проекта

```
lappo_ideal/
├── config.py              # Настройки и конфигурация
├── database.py            # Работа с SQLite
├── google_sheets.py       # Загрузка данных из Google Sheets
├── google_docs.py         # Скачивание шаблонов из Google Docs
├── template_engine.py     # Замена плейсхолдеров в .docx
├── email_sender.py        # Отправка писем через Gmail
├── gui.py                 # Графический интерфейс (tkinter)
├── main.py                # Точка входа
├── requirements.txt       # Зависимости Python
├── .env.example           # Пример файла настроек
├── .gitignore             # Файлы, игнорируемые Git
├── SETUP.md               # Эта инструкция
├── credentials/           # Учётные данные Google API
│   └── credentials.json   # (создаётся вами)
├── templates_cache/       # Скачанные шаблоны
└── output/                # Сгенерированные документы
```

"""
Модуль для работы с Google Sheets API.

Обеспечивает:
- Извлечение ID таблицы из URL
- Авторизацию через OAuth2 или сервисный аккаунт
- Чтение данных из Google Sheets
- Сохранение данных в локальную SQLite базу
"""

import re
from pathlib import Path
from typing import Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

import config
from database import Database


def extract_sheet_id(url: str) -> Optional[str]:
    """
    Извлекает ID таблицы из URL Google Sheets.

    Поддерживает форматы:
    - https://docs.google.com/spreadsheets/d/SHEET_ID/edit
    - https://docs.google.com/spreadsheets/d/SHEET_ID/edit#gid=0
    - https://docs.google.com/spreadsheets/d/SHEET_ID
    """
    pattern = r"spreadsheets/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, url)
    if match:
        return match.group(1)
    return None


def get_google_credentials() -> Credentials:
    """
    Получает учётные данные Google API.

    Использует OAuth2 flow с сохранением токена для повторного использования.
    При первом запуске откроется окно браузера для авторизации.
    """
    creds = None

    # Проверяем сохранённый токен
    if config.GOOGLE_TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(
            str(config.GOOGLE_TOKEN_PATH), config.GOOGLE_SCOPES
        )

    # Если токен недействителен или отсутствует
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                # Токен невозможно обновить — запускаем полную авторизацию
                creds = _run_auth_flow()
        else:
            creds = _run_auth_flow()

        # Сохраняем токен
        config.GOOGLE_TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(config.GOOGLE_TOKEN_PATH, "w") as token_file:
            token_file.write(creds.to_json())

    return creds


def _run_auth_flow() -> Credentials:
    """Запускает OAuth2 авторизацию через браузер."""
    if not config.GOOGLE_CREDENTIALS_PATH.exists():
        raise FileNotFoundError(
            f"Файл учётных данных не найден: {config.GOOGLE_CREDENTIALS_PATH}\n"
            f"Скачайте credentials.json из Google Cloud Console и "
            f"поместите его в папку credentials/"
        )

    flow = InstalledAppFlow.from_client_secrets_file(
        str(config.GOOGLE_CREDENTIALS_PATH), config.GOOGLE_SCOPES
    )
    return flow.run_local_server(port=0)


def fetch_sheet_data(
    sheet_url: str, range_name: str = ""
) -> tuple[str, str, list[str], list[list[str]]]:
    """
    Загружает данные из Google Sheets.

    Аргументы:
        sheet_url: Полный URL таблицы Google Sheets
        range_name: Диапазон ячеек (например, 'Лист1!A:Z').
                    Если пустой — читает первый лист целиком.

    Возвращает:
        Кортеж (sheet_id, sheet_title, заголовки, строки_данных)
    """
    sheet_id = extract_sheet_id(sheet_url)
    if not sheet_id:
        raise ValueError(
            "Не удалось извлечь ID таблицы из URL.\n"
            "Убедитесь, что ссылка имеет формат:\n"
            "https://docs.google.com/spreadsheets/d/ID_ТАБЛИЦЫ/edit"
        )

    creds = get_google_credentials()
    service = build("sheets", "v4", credentials=creds)

    # Получаем информацию о таблице
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=sheet_id
    ).execute()
    sheet_title = spreadsheet.get("properties", {}).get("title", "Без названия")

    # Если диапазон не указан — берём имя первого листа
    if not range_name:
        first_sheet = spreadsheet["sheets"][0]
        sheet_name = first_sheet["properties"]["title"]
        range_name = f"'{sheet_name}'"

    # Читаем данные
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=range_name,
        valueRenderOption="FORMATTED_VALUE"
    ).execute()

    values = result.get("values", [])

    if not values:
        raise ValueError("Таблица пуста — нет данных для загрузки.")

    # Первая строка = заголовки, остальные = данные
    headers = values[0]
    data_rows = values[1:] if len(values) > 1 else []

    return sheet_id, sheet_title, headers, data_rows


def sync_sheet_to_database(
    sheet_url: str, range_name: str = "", db: Optional[Database] = None
) -> tuple[str, int]:
    """
    Загружает данные из Google Sheets и сохраняет в локальную БД.

    Аргументы:
        sheet_url: URL таблицы Google Sheets
        range_name: Диапазон ячеек (опционально)
        db: Экземпляр Database (если None — создаётся новый)

    Возвращает:
        Кортеж (имя_таблицы_в_БД, количество_строк)
    """
    close_db = False
    if db is None:
        db = Database()
        close_db = True

    try:
        sheet_id, sheet_title, headers, data_rows = fetch_sheet_data(
            sheet_url, range_name
        )

        # Создаём таблицу в БД
        table_name = db.create_data_table(
            sheet_id=sheet_id,
            headers=headers,
            sheet_url=sheet_url,
            sheet_name=sheet_title
        )

        # Вставляем данные
        db.insert_rows(table_name, headers, data_rows)

        return table_name, len(data_rows)
    finally:
        if close_db:
            db.close()

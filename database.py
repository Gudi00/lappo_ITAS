"""
Модуль работы с локальной базой данных SQLite.

Обеспечивает:
- Создание системных таблиц (реестр таблиц, шаблонов, лог писем)
- Динамическое создание таблиц на основе заголовков Google Sheets
- CRUD-операции для данных из таблиц
- Хранение метаданных о загруженных листах и шаблонах
"""

import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional

import config


def _sanitize_column_name(name: str) -> str:
    """
    Преобразует имя столбца в безопасное для SQLite.

    Кириллица и другие символы заменяются, но оригинальные имена
    сохраняются в columns_json реестра.
    """
    # Убираем лишние пробелы
    name = name.strip()
    # Заменяем пробелы и спецсимволы на подчёркивания
    sanitized = re.sub(r"[^\w]", "_", name, flags=re.UNICODE)
    # Убираем множественные подчёркивания
    sanitized = re.sub(r"_+", "_", sanitized)
    # Убираем подчёркивания в начале и конце
    sanitized = sanitized.strip("_")
    # Если имя начинается с цифры — добавляем префикс
    if sanitized and sanitized[0].isdigit():
        sanitized = "col_" + sanitized
    # Если после всех преобразований имя пустое
    if not sanitized:
        sanitized = "unnamed_column"
    return sanitized.lower()


def _generate_table_name(sheet_id: str) -> str:
    """Генерирует имя таблицы на основе ID листа Google Sheets."""
    # Берём первые 16 символов ID для краткости
    safe_id = re.sub(r"[^a-zA-Z0-9]", "_", sheet_id[:16])
    return f"sheet_{safe_id}"


class Database:
    """Класс для работы с локальной SQLite базой данных."""

    def __init__(self, db_path: Optional[Path] = None):
        self.db_path = db_path or config.DATABASE_PATH
        self.connection = sqlite3.connect(str(self.db_path))
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA journal_mode=WAL")
        self.connection.execute("PRAGMA encoding='UTF-8'")
        self._create_system_tables()

    def _create_system_tables(self):
        """Создаёт системные таблицы при первом запуске."""
        cursor = self.connection.cursor()

        # Реестр загруженных Google Sheets
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sheets_registry (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sheet_url TEXT NOT NULL,
                sheet_id TEXT NOT NULL UNIQUE,
                sheet_name TEXT,
                table_name TEXT NOT NULL UNIQUE,
                last_synced TIMESTAMP,
                columns_json TEXT,
                row_count INTEGER DEFAULT 0
            )
        """)

        # Реестр шаблонов Google Docs
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS templates_registry (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                doc_url TEXT NOT NULL,
                doc_id TEXT NOT NULL UNIQUE,
                doc_title TEXT,
                local_path TEXT,
                placeholders_json TEXT,
                last_downloaded TIMESTAMP
            )
        """)

        # Журнал отправленных писем
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS email_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                recipient TEXT NOT NULL,
                subject TEXT,
                body_preview TEXT,
                sent_at TIMESTAMP,
                status TEXT DEFAULT 'pending',
                error_message TEXT,
                attachment_path TEXT
            )
        """)

        self.connection.commit()

    # === Работа с данными из Google Sheets ===

    def create_data_table(
        self, sheet_id: str, headers: list[str], sheet_url: str = "",
        sheet_name: str = ""
    ) -> str:
        """
        Создаёт (или пересоздаёт) таблицу для данных из Google Sheets.

        Возвращает имя созданной таблицы.
        """
        table_name = _generate_table_name(sheet_id)

        # Маппинг: оригинальное имя -> безопасное имя
        columns_map = {}
        sanitized_headers = []
        seen = set()
        for header in headers:
            sanitized = _sanitize_column_name(header)
            # Обработка дублирующихся имён
            original_sanitized = sanitized
            counter = 1
            while sanitized in seen:
                sanitized = f"{original_sanitized}_{counter}"
                counter += 1
            seen.add(sanitized)
            columns_map[header] = sanitized
            sanitized_headers.append(sanitized)

        # Удаляем старую таблицу, если есть
        self.connection.execute(f"DROP TABLE IF EXISTS [{table_name}]")

        # Создаём новую таблицу
        columns_sql = ", ".join(
            [f"[{col}] TEXT" for col in sanitized_headers]
        )
        self.connection.execute(
            f"CREATE TABLE [{table_name}] ("
            f"  id INTEGER PRIMARY KEY AUTOINCREMENT, "
            f"  {columns_sql}"
            f")"
        )

        # Обновляем реестр
        self.connection.execute("""
            INSERT INTO sheets_registry
                (sheet_url, sheet_id, sheet_name, table_name,
                 last_synced, columns_json)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(sheet_id) DO UPDATE SET
                sheet_url = excluded.sheet_url,
                sheet_name = excluded.sheet_name,
                table_name = excluded.table_name,
                last_synced = excluded.last_synced,
                columns_json = excluded.columns_json
        """, (
            sheet_url, sheet_id, sheet_name, table_name,
            datetime.now().isoformat(),
            json.dumps(columns_map, ensure_ascii=False)
        ))

        self.connection.commit()
        return table_name

    def insert_rows(self, table_name: str, headers: list[str],
                    rows: list[list[str]]):
        """Вставляет строки данных в таблицу."""
        if not rows:
            return

        # Получаем маппинг столбцов из реестра
        cursor = self.connection.execute(
            "SELECT columns_json FROM sheets_registry WHERE table_name = ?",
            (table_name,)
        )
        row = cursor.fetchone()
        if not row:
            raise ValueError(f"Таблица {table_name} не найдена в реестре")

        columns_map = json.loads(row["columns_json"])
        sanitized_headers = [columns_map.get(h, _sanitize_column_name(h))
                             for h in headers]

        placeholders = ", ".join(["?"] * len(sanitized_headers))
        columns = ", ".join([f"[{col}]" for col in sanitized_headers])

        for data_row in rows:
            # Дополняем строку пустыми значениями, если столбцов не хватает
            padded = list(data_row) + [""] * (len(sanitized_headers) - len(data_row))
            padded = padded[:len(sanitized_headers)]
            self.connection.execute(
                f"INSERT INTO [{table_name}] ({columns}) VALUES ({placeholders})",
                padded
            )

        # Обновляем количество строк в реестре
        self.connection.execute(
            "UPDATE sheets_registry SET row_count = ? WHERE table_name = ?",
            (len(rows), table_name)
        )
        self.connection.commit()

    def get_all_data(self, table_name: str) -> tuple[list[str], list[list[str]]]:
        """
        Возвращает все данные из таблицы.

        Возвращает кортеж: (список заголовков, список строк).
        """
        cursor = self.connection.execute(f"SELECT * FROM [{table_name}]")
        columns = [desc[0] for desc in cursor.description if desc[0] != "id"]

        rows = []
        for row in cursor.fetchall():
            rows.append([row[col] for col in columns])

        return columns, rows

    def get_original_headers(self, table_name: str) -> dict[str, str]:
        """
        Возвращает маппинг оригинальных заголовков к безопасным именам.

        Ключ = оригинальное имя, значение = имя в SQLite.
        """
        cursor = self.connection.execute(
            "SELECT columns_json FROM sheets_registry WHERE table_name = ?",
            (table_name,)
        )
        row = cursor.fetchone()
        if row and row["columns_json"]:
            return json.loads(row["columns_json"])
        return {}

    def get_row_by_index(self, table_name: str, index: int) -> Optional[dict]:
        """Возвращает строку по индексу (начиная с 0)."""
        cursor = self.connection.execute(
            f"SELECT * FROM [{table_name}] LIMIT 1 OFFSET ?", (index,)
        )
        row = cursor.fetchone()
        if row:
            return {key: row[key] for key in row.keys() if key != "id"}
        return None

    def get_row_count(self, table_name: str) -> int:
        """Возвращает количество строк в таблице."""
        cursor = self.connection.execute(
            f"SELECT COUNT(*) as cnt FROM [{table_name}]"
        )
        return cursor.fetchone()["cnt"]

    # === Работа с реестром таблиц ===

    def get_all_sheets(self) -> list[dict]:
        """Возвращает список всех загруженных таблиц."""
        cursor = self.connection.execute(
            "SELECT * FROM sheets_registry ORDER BY last_synced DESC"
        )
        return [dict(row) for row in cursor.fetchall()]

    def get_sheet_by_id(self, sheet_id: str) -> Optional[dict]:
        """Находит запись о листе по его Google ID."""
        cursor = self.connection.execute(
            "SELECT * FROM sheets_registry WHERE sheet_id = ?", (sheet_id,)
        )
        row = cursor.fetchone()
        return dict(row) if row else None

    # === Работа с реестром шаблонов ===

    def register_template(self, doc_url: str, doc_id: str,
                          doc_title: str, local_path: str,
                          placeholders: list[str]):
        """Регистрирует скачанный шаблон в реестре."""
        self.connection.execute("""
            INSERT INTO templates_registry
                (doc_url, doc_id, doc_title, local_path,
                 placeholders_json, last_downloaded)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(doc_id) DO UPDATE SET
                doc_url = excluded.doc_url,
                doc_title = excluded.doc_title,
                local_path = excluded.local_path,
                placeholders_json = excluded.placeholders_json,
                last_downloaded = excluded.last_downloaded
        """, (
            doc_url, doc_id, doc_title, str(local_path),
            json.dumps(placeholders, ensure_ascii=False),
            datetime.now().isoformat()
        ))
        self.connection.commit()

    def get_all_templates(self) -> list[dict]:
        """Возвращает список всех зарегистрированных шаблонов."""
        cursor = self.connection.execute(
            "SELECT * FROM templates_registry ORDER BY last_downloaded DESC"
        )
        return [dict(row) for row in cursor.fetchall()]

    def get_template_by_id(self, doc_id: str) -> Optional[dict]:
        """Находит шаблон по Google Doc ID."""
        cursor = self.connection.execute(
            "SELECT * FROM templates_registry WHERE doc_id = ?", (doc_id,)
        )
        row = cursor.fetchone()
        return dict(row) if row else None

    # === Журнал писем ===

    def log_email(self, recipient: str, subject: str,
                  body_preview: str = "", status: str = "pending",
                  error_message: str = "",
                  attachment_path: str = "") -> int:
        """Записывает информацию об отправленном письме."""
        cursor = self.connection.execute("""
            INSERT INTO email_log
                (recipient, subject, body_preview, sent_at,
                 status, error_message, attachment_path)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            recipient, subject, body_preview,
            datetime.now().isoformat(), status,
            error_message, attachment_path
        ))
        self.connection.commit()
        return cursor.lastrowid

    def update_email_status(self, log_id: int, status: str,
                            error_message: str = ""):
        """Обновляет статус отправки письма."""
        self.connection.execute("""
            UPDATE email_log SET status = ?, error_message = ?, sent_at = ?
            WHERE id = ?
        """, (status, error_message, datetime.now().isoformat(), log_id))
        self.connection.commit()

    def get_email_log(self, limit: int = 100) -> list[dict]:
        """Возвращает журнал последних отправленных писем."""
        cursor = self.connection.execute(
            "SELECT * FROM email_log ORDER BY sent_at DESC LIMIT ?",
            (limit,)
        )
        return [dict(row) for row in cursor.fetchall()]

    # === Общие методы ===

    def get_table_names(self) -> list[str]:
        """Возвращает список всех пользовательских таблиц (sheet_*)."""
        cursor = self.connection.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name LIKE 'sheet_%'
            ORDER BY name
        """)
        return [row["name"] for row in cursor.fetchall()]

    def close(self):
        """Закрывает соединение с базой данных."""
        self.connection.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

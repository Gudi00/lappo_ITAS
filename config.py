"""
Конфигурация приложения.

Загружает настройки из файла .env и предоставляет
централизованный доступ ко всем параметрам.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Корневая директория проекта
BASE_DIR = Path(__file__).resolve().parent

# Загружаем переменные из .env
load_dotenv(BASE_DIR / ".env")

# === Google API ===
GOOGLE_CREDENTIALS_PATH = Path(
    os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials/credentials.json")
)
if not GOOGLE_CREDENTIALS_PATH.is_absolute():
    GOOGLE_CREDENTIALS_PATH = BASE_DIR / GOOGLE_CREDENTIALS_PATH

GOOGLE_TOKEN_PATH = BASE_DIR / "credentials" / "token.json"

# Области доступа Google API (только чтение таблиц + экспорт документов)
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# === Пути к папкам ===
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "output"))
if not OUTPUT_DIR.is_absolute():
    OUTPUT_DIR = BASE_DIR / OUTPUT_DIR

TEMPLATES_CACHE_DIR = Path(os.getenv("TEMPLATES_CACHE_DIR", "templates_cache"))
if not TEMPLATES_CACHE_DIR.is_absolute():
    TEMPLATES_CACHE_DIR = BASE_DIR / TEMPLATES_CACHE_DIR

# Создаём папки, если их нет
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_CACHE_DIR.mkdir(parents=True, exist_ok=True)

# === База данных ===
DATABASE_PATH = Path(os.getenv("DATABASE_PATH", "app_database.db"))
if not DATABASE_PATH.is_absolute():
    DATABASE_PATH = BASE_DIR / DATABASE_PATH

# === Gmail SMTP ===
GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS", "")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

"""
Точка входа в приложение «Автоматизация документов».

Запуск:
    python main.py
"""

import sys
from pathlib import Path


def check_dependencies():
    """Проверяет наличие необходимых библиотек."""
    missing = []
    try:
        import google.auth  # noqa: F401
    except ImportError:
        missing.append("google-auth")
    try:
        import googleapiclient  # noqa: F401
    except ImportError:
        missing.append("google-api-python-client")
    try:
        import google_auth_oauthlib  # noqa: F401
    except ImportError:
        missing.append("google-auth-oauthlib")
    try:
        import docx  # noqa: F401
    except ImportError:
        missing.append("python-docx")
    try:
        import dotenv  # noqa: F401
    except ImportError:
        missing.append("python-dotenv")

    if missing:
        print("Не установлены необходимые библиотеки:")
        for lib in missing:
            print(f"  - {lib}")
        print()
        print("Установите их командой:")
        print(f"  pip install {' '.join(missing)}")
        print()
        print("Или установите все зависимости:")
        print("  pip install -r requirements.txt")
        sys.exit(1)


def check_credentials():
    """Проверяет наличие файла учётных данных Google API."""
    from config import GOOGLE_CREDENTIALS_PATH

    if not GOOGLE_CREDENTIALS_PATH.exists():
        print("=" * 60)
        print("ПЕРВЫЙ ЗАПУСК — НАСТРОЙКА GOOGLE API")
        print("=" * 60)
        print()
        print("Файл учётных данных не найден:")
        print(f"  {GOOGLE_CREDENTIALS_PATH}")
        print()
        print("Для работы приложения необходимо:")
        print()
        print("1. Создайте проект в Google Cloud Console:")
        print("   https://console.cloud.google.com/")
        print()
        print("2. Включите API:")
        print("   - Google Sheets API")
        print("   - Google Drive API")
        print()
        print("3. Создайте учётные данные OAuth 2.0:")
        print("   - Тип: Desktop application")
        print("   - Скачайте JSON файл")
        print()
        print("4. Переименуйте скачанный файл в 'credentials.json'")
        print(f"   и поместите в: {GOOGLE_CREDENTIALS_PATH.parent}/")
        print()
        print("5. Скопируйте .env.example в .env и заполните настройки:")
        print("   cp .env.example .env")
        print()
        print("=" * 60)
        print()

        answer = input("Продолжить без Google API? (да/нет): ").strip().lower()
        if answer not in ("да", "д", "y", "yes"):
            sys.exit(0)


def main():
    """Главная функция запуска приложения."""
    print("Автоматизация документов — запуск...")
    print()

    # Проверяем зависимости
    check_dependencies()

    # Проверяем учётные данные
    check_credentials()

    # Запускаем GUI
    from gui import App

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

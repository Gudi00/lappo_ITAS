"""
Модуль для работы с Google Docs API / Drive API.

Обеспечивает:
- Извлечение ID документа из URL
- Скачивание Google Docs как .docx файлы
- Регистрацию шаблонов в локальной базе данных
"""

import re
from pathlib import Path
from typing import Optional

from googleapiclient.discovery import build

import config
from database import Database
from google_sheets import get_google_credentials


def extract_doc_id(url: str) -> Optional[str]:
    """
    Извлекает ID документа из URL Google Docs.

    Поддерживает форматы:
    - https://docs.google.com/document/d/DOC_ID/edit
    - https://docs.google.com/document/d/DOC_ID/
    - https://docs.google.com/document/d/DOC_ID
    """
    pattern = r"document/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, url)
    if match:
        return match.group(1)
    return None


def download_template(
    doc_url: str, db: Optional[Database] = None
) -> Path:
    """
    Скачивает Google Doc как .docx файл в локальную папку.

    Аргументы:
        doc_url: Полный URL Google Docs
        db: Экземпляр Database для регистрации шаблона

    Возвращает:
        Path к скачанному .docx файлу
    """
    doc_id = extract_doc_id(doc_url)
    if not doc_id:
        raise ValueError(
            "Не удалось извлечь ID документа из URL.\n"
            "Убедитесь, что ссылка имеет формат:\n"
            "https://docs.google.com/document/d/ID_ДОКУМЕНТА/edit"
        )

    creds = get_google_credentials()

    # Получаем метаданные документа через Drive API
    drive_service = build("drive", "v3", credentials=creds)

    file_metadata = drive_service.files().get(
        fileId=doc_id, fields="name,mimeType"
    ).execute()
    doc_title = file_metadata.get("name", "template")
    mime_type = file_metadata.get("mimeType", "")

    # Формируем безопасное имя файла
    safe_title = re.sub(r'[<>:"/\\|?*]', "_", doc_title)
    if not safe_title.endswith(".docx"):
        safe_title += ".docx"
    local_path = config.TEMPLATES_CACHE_DIR / safe_title

    # Нативный Google Doc — экспортируем; загруженный .docx — скачиваем напрямую
    if mime_type == "application/vnd.google-apps.document":
        request = drive_service.files().export_media(
            fileId=doc_id,
            mimeType="application/vnd.openxmlformats-officedocument"
                     ".wordprocessingml.document"
        )
    else:
        request = drive_service.files().get_media(fileId=doc_id)

    with open(local_path, "wb") as f:
        f.write(request.execute())

    # Находим плейсхолдеры в скачанном файле
    from template_engine import find_placeholders
    placeholders = find_placeholders(local_path)

    # Регистрируем шаблон в базе данных
    if db is not None:
        db.register_template(
            doc_url=doc_url,
            doc_id=doc_id,
            doc_title=doc_title,
            local_path=str(local_path),
            placeholders=placeholders
        )

    return local_path

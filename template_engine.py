"""
Движок обработки шаблонов .docx.

Обеспечивает:
- Поиск плейсхолдеров ${...} в документе
- Замену ${column_name} на значения из базы данных
- Вставку таблиц по плейсхолдеру ${table}
- Генерацию одного документа на каждую строку данных
"""

import copy
import re
from pathlib import Path
from typing import Optional

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

import config
from database import Database

# Регулярное выражение для поиска плейсхолдеров ${...}
PLACEHOLDER_PATTERN = re.compile(r"\$\{([^}]+)\}")


def find_placeholders(docx_path: Path) -> list[str]:
    """
    Находит все плейсхолдеры ${...} в документе .docx.

    Ищет в:
    - Параграфах основного текста
    - Таблицах
    - Колонтитулах (headers/footers)

    Возвращает список уникальных имён плейсхолдеров (без ${}).
    """
    doc = Document(str(docx_path))
    placeholders = set()

    # Поиск в параграфах
    for paragraph in doc.paragraphs:
        text = paragraph.text
        for match in PLACEHOLDER_PATTERN.finditer(text):
            placeholders.add(match.group(1))

    # Поиск в таблицах документа
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for match in PLACEHOLDER_PATTERN.finditer(paragraph.text):
                        placeholders.add(match.group(1))

    # Поиск в колонтитулах
    for section in doc.sections:
        for header_footer in [section.header, section.footer]:
            if header_footer is not None:
                for paragraph in header_footer.paragraphs:
                    for match in PLACEHOLDER_PATTERN.finditer(paragraph.text):
                        placeholders.add(match.group(1))

    return sorted(placeholders)


def _replace_text_in_paragraph(paragraph, replacements: dict[str, str]):
    """
    Заменяет плейсхолдеры в параграфе, сохраняя форматирование.

    Работает на уровне runs, чтобы не терять стили шрифта.
    Иногда Google Docs разбивает ${placeholder} на несколько runs,
    поэтому сначала собираем полный текст, а потом перестраиваем runs.
    """
    # Собираем полный текст параграфа
    full_text = paragraph.text

    # Проверяем, есть ли хоть один плейсхолдер
    if not PLACEHOLDER_PATTERN.search(full_text):
        return

    # Выполняем все замены
    new_text = full_text
    for placeholder, value in replacements.items():
        new_text = new_text.replace(f"${{{placeholder}}}", str(value))

    # Если текст не изменился — ничего не делаем
    if new_text == full_text:
        return

    # Сохраняем форматирование первого run
    if paragraph.runs:
        # Сохраняем стиль первого run
        first_run = paragraph.runs[0]
        font_props = {}
        if first_run.font.name:
            font_props["name"] = first_run.font.name
        if first_run.font.size:
            font_props["size"] = first_run.font.size
        if first_run.font.bold is not None:
            font_props["bold"] = first_run.font.bold
        if first_run.font.italic is not None:
            font_props["italic"] = first_run.font.italic

        # Очищаем все runs
        for run in paragraph.runs:
            run.text = ""

        # Устанавливаем текст в первый run
        paragraph.runs[0].text = new_text

        # Восстанавливаем форматирование
        for prop, value in font_props.items():
            setattr(paragraph.runs[0].font, prop, value)
    else:
        # Если runs нет — просто добавляем
        paragraph.add_run(new_text)


def _insert_table_after_paragraph(doc, paragraph, headers: list[str],
                                  rows: list[list[str]]):
    """
    Вставляет таблицу с данными после указанного параграфа.

    Удаляет параграф с ${table} и вставляет вместо него таблицу.
    """
    # Создаём таблицу
    num_cols = len(headers)
    num_rows = len(rows) + 1  # +1 для заголовков

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Применяем стиль таблицы
    try:
        table.style = "Table Grid"
    except KeyError:
        pass  # Стиль может отсутствовать

    # Заполняем заголовки
    header_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = str(header)
        # Делаем заголовки жирными
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(10)

    # Заполняем данные
    for row_idx, row_data in enumerate(rows):
        table_row = table.rows[row_idx + 1]
        for col_idx, value in enumerate(row_data):
            if col_idx < num_cols:
                cell = table_row.cells[col_idx]
                cell.text = str(value) if value else ""
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(10)

    # Перемещаем таблицу на место параграфа с ${table}
    paragraph_element = paragraph._element
    table_element = table._tbl

    # Вставляем таблицу перед параграфом
    paragraph_element.addprevious(table_element)

    # Удаляем параграф с ${table}
    parent = paragraph_element.getparent()
    parent.remove(paragraph_element)


def generate_document(
    template_path: Path,
    row_data: dict[str, str],
    output_filename: str,
    table_data: Optional[tuple[list[str], list[list[str]]]] = None,
    original_headers_map: Optional[dict[str, str]] = None,
) -> Path:
    """
    Генерирует документ из шаблона, заменяя плейсхолдеры.

    Аргументы:
        template_path: Путь к файлу шаблона .docx
        row_data: Словарь {имя_столбца: значение} для текущей строки
        output_filename: Имя выходного файла
        table_data: Кортеж (заголовки, строки) для вставки таблицы
        original_headers_map: Маппинг оригинальных имён к безопасным

    Возвращает:
        Path к сгенерированному файлу
    """
    doc = Document(str(template_path))

    # Строим замены: поддерживаем и оригинальные, и безопасные имена
    replacements = {}
    if original_headers_map:
        # Маппинг в обе стороны
        reverse_map = {v: k for k, v in original_headers_map.items()}
        for orig_name, safe_name in original_headers_map.items():
            value = row_data.get(safe_name, "")
            replacements[orig_name] = value  # ${Оригинальное имя}
            replacements[safe_name] = value  # ${безопасное_имя}
    else:
        replacements = dict(row_data)

    # Обработка параграфов
    paragraphs_to_replace_with_table = []
    for paragraph in doc.paragraphs:
        if "${table}" in paragraph.text.lower() or "${TABLE}" in paragraph.text:
            paragraphs_to_replace_with_table.append(paragraph)
        else:
            _replace_text_in_paragraph(paragraph, replacements)

    # Обработка таблиц в документе
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_text_in_paragraph(paragraph, replacements)

    # Обработка колонтитулов
    for section in doc.sections:
        for header_footer in [section.header, section.footer]:
            if header_footer is not None:
                for paragraph in header_footer.paragraphs:
                    _replace_text_in_paragraph(paragraph, replacements)

    # Вставка таблиц (${table})
    if table_data and paragraphs_to_replace_with_table:
        headers, rows = table_data
        # Используем оригинальные имена для заголовков таблицы
        display_headers = headers
        if original_headers_map:
            reverse_map = {v: k for k, v in original_headers_map.items()}
            display_headers = [reverse_map.get(h, h) for h in headers]

        for paragraph in paragraphs_to_replace_with_table:
            _insert_table_after_paragraph(
                doc, paragraph, display_headers, rows
            )

    # Сохраняем результат
    output_path = config.OUTPUT_DIR / output_filename
    doc.save(str(output_path))

    return output_path


def generate_documents_for_all_rows(
    template_path: Path,
    table_name: str,
    filename_column: Optional[str] = None,
    include_table: bool = False,
    db: Optional[Database] = None
) -> list[Path]:
    """
    Генерирует по одному документу на каждую строку данных.

    Аргументы:
        template_path: Путь к шаблону .docx
        table_name: Имя таблицы в SQLite
        filename_column: Столбец для имени файла (если None — используется номер)
        include_table: Вставлять ли полную таблицу по ${table}
        db: Экземпляр Database

    Возвращает:
        Список путей к сгенерированным файлам
    """
    close_db = False
    if db is None:
        db = Database()
        close_db = True

    try:
        headers, all_rows = db.get_all_data(table_name)
        original_map = db.get_original_headers(table_name)

        # Данные для таблицы (если нужна)
        table_data = (headers, all_rows) if include_table else None

        generated_files = []

        for idx, row in enumerate(all_rows):
            # Создаём словарь данных для строки
            row_data = dict(zip(headers, row))

            # Формируем имя файла
            if filename_column and filename_column in row_data:
                safe_name = re.sub(r'[<>:"/\\|?*]', "_",
                                   str(row_data[filename_column]))
                output_name = f"{safe_name}.docx"
            else:
                output_name = f"document_{idx + 1:03d}.docx"

            output_path = generate_document(
                template_path=template_path,
                row_data=row_data,
                output_filename=output_name,
                table_data=table_data,
                original_headers_map=original_map
            )
            generated_files.append(output_path)

        return generated_files
    finally:
        if close_db:
            db.close()

"""
Движок обработки шаблонов .docx.

Синтаксис плейсхолдеров:
  ${ФИО}                   — значение поля текущей строки (студента)
  ${table}                 — таблица всех строк, все столбцы
  ${table:ФИО,Группа}      — таблица всех строк, только указанные столбцы
  ${row}                   — текущая строка как таблица, все столбцы
  ${row:ФИО,Группа}        — текущая строка как таблица, только указанные столбцы
"""

import re
from pathlib import Path
from typing import Optional

from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_TABLE_ALIGNMENT

import config
from database import Database

# Регулярное выражение для поиска плейсхолдеров ${...}
PLACEHOLDER_PATTERN = re.compile(r"\$\{([^}]+)\}")

# Плейсхолдеры таблиц: ${table}, ${table:col1,col2}, ${row}, ${row:col1,col2}
TABLE_PLACEHOLDER_PATTERN = re.compile(
    r"\$\{(table|row)(?::([^}]*))?\}", re.IGNORECASE
)


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


def _filter_columns(
    headers: list[str],
    rows: list[list[str]],
    columns: Optional[list[str]],
    original_headers_map: Optional[dict[str, str]] = None,
) -> tuple[list[str], list[list[str]]]:
    """
    Фильтрует столбцы по списку columns.
    Если columns=None — возвращает все столбцы без изменений.
    Принимает как оригинальные (русские), так и безопасные имена.
    """
    if not columns:
        return headers, rows

    # Строим индекс: lowercase имя → индекс (по оригинальным именам)
    col_index: dict[str, int] = {}
    for i, h in enumerate(headers):
        col_index[h.lower()] = i
    # Добавляем поиск и по безопасным (санированным) именам
    if original_headers_map:
        for orig, safe in original_headers_map.items():
            idx = col_index.get(orig.lower(), -1)
            if idx != -1:
                col_index[safe.lower()] = idx

    selected_indices = []
    selected_headers = []
    for col in columns:
        idx = col_index.get(col.strip().lower(), -1)
        if idx != -1 and idx not in selected_indices:
            selected_indices.append(idx)
            selected_headers.append(headers[idx])

    if not selected_indices:
        return headers, rows

    filtered_rows = [[row[i] for i in selected_indices if i < len(row)]
                     for row in rows]
    return selected_headers, filtered_rows


def _insert_table_after_paragraph(
    doc,
    paragraph,
    headers: list[str],
    rows: list[list[str]],
):
    """Вставляет таблицу с данными вместо параграфа-плейсхолдера."""
    num_cols = len(headers)
    num_rows = len(rows) + 1  # +1 для заголовков

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    try:
        table.style = "Table Grid"
    except KeyError:
        pass

    # Заголовки
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = str(header)
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)

    # Данные
    for row_idx, row_data in enumerate(rows):
        table_row = table.rows[row_idx + 1]
        for col_idx, value in enumerate(row_data):
            if col_idx < num_cols:
                cell = table_row.cells[col_idx]
                cell.text = str(value) if value else ""
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(10)

    # Вставляем таблицу перед параграфом-плейсхолдером, затем удаляем его
    paragraph._element.addprevious(table._tbl)
    paragraph._element.getparent().remove(paragraph._element)


def generate_document(
    template_path: Path,
    row_data: dict[str, str],
    output_filename: str,
    all_rows_data: Optional[tuple[list[str], list[list[str]]]] = None,
    original_headers_map: Optional[dict[str, str]] = None,
) -> Path:
    """
    Генерирует документ из шаблона, заменяя плейсхолдеры.

    Поддерживаемые плейсхолдеры:
      ${ФИО}               — значение поля текущей строки
      ${table}             — таблица всех строк, все столбцы
      ${table:ФИО,Группа}  — таблица всех строк, только указанные столбцы
      ${row}               — текущая строка как таблица, все столбцы
      ${row:ФИО,Группа}    — текущая строка как таблица, только указанные столбцы
    """
    doc = Document(str(template_path))

    # Строим замены для обычных полей (оригинальные + безопасные имена)
    replacements: dict[str, str] = {}
    if original_headers_map:
        for orig_name, safe_name in original_headers_map.items():
            value = row_data.get(safe_name, "")
            replacements[orig_name] = value
            replacements[safe_name] = value
    else:
        replacements = dict(row_data)

    # Подготавливаем данные текущей строки как таблицу (для ${row})
    if original_headers_map:
        row_headers = list(original_headers_map.keys())
        row_values = [row_data.get(safe, "") for safe in original_headers_map.values()]
    else:
        row_headers = list(row_data.keys())
        row_values = list(row_data.values())
    current_row_as_table = (row_headers, [row_values])

    # Данные всех строк для ${table}
    if all_rows_data:
        all_headers_raw, all_rows_raw = all_rows_data
        if original_headers_map:
            rev = {v: k for k, v in original_headers_map.items()}
            display_all_headers = [rev.get(h, h) for h in all_headers_raw]
        else:
            display_all_headers = all_headers_raw
        all_table_data = (display_all_headers, all_rows_raw)
    else:
        all_table_data = None

    def _process_paragraphs(paragraphs):
        """Обрабатывает список параграфов: заменяет поля, собирает таблицы."""
        pending_tables = []
        for paragraph in paragraphs:
            text = paragraph.text
            m = TABLE_PLACEHOLDER_PATTERN.search(text)
            if m:
                kind = m.group(1).lower()       # "table" или "row"
                cols_str = m.group(2)           # "ФИО,Группа" или None
                columns = [c.strip() for c in cols_str.split(",")] if cols_str else None
                pending_tables.append((paragraph, kind, columns))
            else:
                _replace_text_in_paragraph(paragraph, replacements)
        return pending_tables

    # Обрабатываем основные параграфы
    pending = _process_paragraphs(doc.paragraphs)

    # Обрабатываем ячейки таблиц документа
    for tbl in doc.tables:
        for trow in tbl.rows:
            for cell in trow.cells:
                pending += _process_paragraphs(cell.paragraphs)

    # Обрабатываем колонтитулы
    for section in doc.sections:
        for hf in [section.header, section.footer]:
            if hf is not None:
                pending += _process_paragraphs(hf.paragraphs)

    # Вставляем таблицы (в обратном порядке, чтобы не сбить позиции)
    for paragraph, kind, columns in reversed(pending):
        if kind == "table":
            if all_table_data is None:
                paragraph._element.getparent().remove(paragraph._element)
                continue
            h, r = _filter_columns(
                all_table_data[0], all_table_data[1], columns, original_headers_map
            )
        else:  # row
            h, r = _filter_columns(
                current_row_as_table[0], current_row_as_table[1],
                columns, original_headers_map
            )
        _insert_table_after_paragraph(doc, paragraph, h, r)

    output_path = config.OUTPUT_DIR / output_filename
    doc.save(str(output_path))
    return output_path


def generate_documents_for_all_rows(
    template_path: Path,
    table_name: str,
    filename_column: Optional[str] = None,
    db: Optional[Database] = None,
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

        # Данные всех строк передаём всегда — generate_document сам решит
        # что вставлять по ${table} и ${row}
        all_rows_data = (headers, all_rows)

        generated_files = []

        for idx, row in enumerate(all_rows):
            row_data = dict(zip(headers, row))

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
                all_rows_data=all_rows_data,
                original_headers_map=original_map,
            )
            generated_files.append(output_path)

        return generated_files
    finally:
        if close_db:
            db.close()

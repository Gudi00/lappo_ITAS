"""
Движок обработки шаблонов .docx.

═══════════════════════════════════════════════════════════════
СИНТАКСИС ШАБЛОНА
═══════════════════════════════════════════════════════════════

1. ТЕКСТОВЫЕ ПОЛЯ (для текущего студента):
   {{ ФИО }}              — значение поля «ФИО»
   {{ Группа }}           — значение поля «Группа»

2. ТАБЛИЦА ВСЕХ СТУДЕНТОВ (через цикл в строке таблицы):
   В строке таблицы разместите:
   {% for s in students %}{{ s.ФИО }} {{ s.Группа }}{% endfor %}

   Строка будет продублирована для каждого студента.
"""

import copy
import re
import shutil
import tempfile
from pathlib import Path
from typing import Optional

from docx import Document
from docxtpl import DocxTemplate

import config
from database import Database

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Паттерн для поиска Jinja2-переменных: {{ имя }}
_VAR_PATTERN = re.compile(r"\{\{\s*([^%{}\s][^{}]*?)\s*\}\}")


def find_placeholders(docx_path: Path) -> list[str]:
    """Находит все плейсхолдеры {{ ... }} в документе."""
    doc = Document(str(docx_path))
    placeholders: set[str] = set()

    def _scan(text: str) -> None:
        for m in _VAR_PATTERN.finditer(text):
            name = m.group(1).strip()
            if name.startswith(("%", "tr ", "for ", "if ", "end")):
                continue
            placeholders.add(name)

    for para in doc.paragraphs:
        _scan(para.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _scan(para.text)

    for section in doc.sections:
        for hf in (section.header, section.footer):
            if hf is not None:
                for para in hf.paragraphs:
                    _scan(para.text)

    return sorted(placeholders)


def _build_student_dict(
    row_data: dict[str, str],
    original_headers_map: Optional[dict[str, str]],
) -> dict[str, str]:
    """Строит словарь полей студента с оригинальными и санированными именами."""
    student: dict[str, str] = {}
    if original_headers_map:
        for orig, safe in original_headers_map.items():
            value = str(row_data.get(safe, ""))
            student[orig] = value
            student[safe] = value
    else:
        student = {k: str(v) for k, v in row_data.items()}
    return student


def _flatten_cell_to_single_paragraph(cell_elem) -> None:
    """
    Объединяет все параграфы ячейки в один параграф с одним текстовым run.
    Используется для строк с циклами, чтобы regex корректно работал с Jinja-тегами,
    разбитыми переносами строк внутри ячейки.
    """
    paragraphs = cell_elem.findall(f'{W_NS}p')
    if not paragraphs:
        return

    # Собираем весь текст из всех параграфов, разделяя их переводом строки
    parts = []
    for p in paragraphs:
        texts = p.findall(f'.//{W_NS}t')
        para_text = ''.join(t.text or '' for t in texts)
        parts.append(para_text)
    full_text = '\n'.join(parts)

    # Оставляем первый параграф, объединяя все run'ы в один
    first_p = paragraphs[0]
    texts = first_p.findall(f'.//{W_NS}t')
    if texts:
        texts[0].text = full_text
        texts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        for t in texts[1:]:
            t.text = ''
    else:
        # Нет существующих <w:t> — создаём новый run
        from docx.oxml import OxmlElement
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.text = full_text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t)
        first_p.append(r)

    # Удаляем остальные параграфы
    for p in paragraphs[1:]:
        cell_elem.remove(p)


def _find_loop_row(table) -> Optional[int]:
    """Находит индекс строки, содержащей {% for s in students %}."""
    for row_idx, row in enumerate(table.rows):
        row_text = '\n'.join(cell.text for cell in row.cells)
        if re.search(r'\{%\s*(?:tr\s+)?for\s+\w+\s+in\s+\w+', row_text):
            return row_idx
    return None


def _substitute_student_in_row(row_element, student: dict[str, str]) -> None:
    """
    Подставляет значения студента в строку таблицы:
    - {{ s.FieldName }} → фактическое значение
    - Удаляет {% for %} и {% endfor %} markers

    Сначала сплющивает содержимое каждой ячейки в один параграф, чтобы
    regex корректно работал с Jinja-тегами, разбитыми на несколько параграфов.
    """
    def _replace(match, student_dict):
        field_name = match.group(1)
        # Убираем ведущие/конечные пробелы и переносы (но сохраняем внутренние)
        field_name = field_name.strip()
        if field_name in student_dict:
            return student_dict[field_name]
        # Пробуем санированную версию
        sanitized = re.sub(r'\s+', '_', field_name).lower()
        if sanitized in student_dict:
            return student_dict[sanitized]
        return ''

    # Обрабатываем каждую ячейку
    for tc in row_element.findall(f'.//{W_NS}tc'):
        # Сплющиваем ячейку в один параграф
        _flatten_cell_to_single_paragraph(tc)

        # Теперь в ячейке один параграф с одним текстовым run — можем делать regex
        for t in tc.findall(f'.//{W_NS}t'):
            if not t.text:
                continue
            text = t.text

            # Удаляем loop markers (с учётом возможных переносов)
            text = re.sub(r'\{%\s*(?:tr\s+)?for[^%]*?%\}', '', text, flags=re.DOTALL)
            text = re.sub(r'\{%\s*(?:tr\s+)?endfor[^%]*?%\}', '', text, flags=re.DOTALL)

            # Заменяем {{ s.FieldName }} на фактическое значение
            # DOTALL чтобы . матчил переносы (для полей с multi-line именами)
            text = re.sub(
                r'\{\{\s*s\.(.*?)\}\}',
                lambda m: _replace(m, student),
                text,
                flags=re.DOTALL,
            )

            t.text = text


def _expand_table_loops(template_path: Path, students_list: list[dict]) -> None:
    """
    Разворачивает циклы в таблицах:
    - Находит строку с {% for s in students %}
    - Дублирует её для каждого студента, подставляя значения
    - Удаляет loop markers
    """
    doc = Document(str(template_path))

    for table in doc.tables:
        loop_row_idx = _find_loop_row(table)
        if loop_row_idx is None:
            continue

        template_row = table.rows[loop_row_idx]
        tr_element = template_row._element
        parent = tr_element.getparent()
        insert_pos = parent.index(tr_element)

        # Генерируем строки для каждого студента (новые — вставляем ПЕРЕД шаблонной)
        new_rows = []
        for student in students_list:
            tr_copy = copy.deepcopy(tr_element)
            _substitute_student_in_row(tr_copy, student)
            new_rows.append(tr_copy)

        # Вставляем новые строки перед шаблонной строкой
        for i, new_row in enumerate(new_rows):
            parent.insert(insert_pos + i, new_row)

        # Удаляем исходную шаблонную строку
        parent.remove(tr_element)

    doc.save(str(template_path))


def generate_document(
    template_path: Path,
    row_data: dict[str, str],
    output_filename: str,
    all_rows_data: Optional[tuple[list[str], list[list]]] = None,
    original_headers_map: Optional[dict[str, str]] = None,
) -> Path:
    """Генерирует один документ из шаблона."""
    # Текущий студент (для прямых плейсхолдеров вне таблиц)
    current_student = _build_student_dict(row_data, original_headers_map)

    # Все студенты (для таблиц с циклами)
    students_list: list[dict[str, str]] = []
    if all_rows_data:
        headers_raw, rows_raw = all_rows_data
        rev = {}
        if original_headers_map:
            rev = {safe: orig for orig, safe in original_headers_map.items()}

        for row in rows_raw:
            s: dict[str, str] = {}
            for h_raw, val in zip(headers_raw, row):
                val_str = str(val) if val is not None else ""
                orig = rev.get(h_raw, h_raw)
                s[orig] = val_str
                s[h_raw] = val_str
            students_list.append(s)

    # Работаем с копией шаблона
    with tempfile.TemporaryDirectory() as tmpdir:
        work_template = Path(tmpdir) / template_path.name
        shutil.copy2(template_path, work_template)

        # Шаг 1: разворачиваем циклы в таблицах вручную (с прямой подстановкой)
        _expand_table_loops(work_template, students_list)

        # Шаг 2: рендерим оставшиеся плейсхолдеры через docxtpl (нетабличные)
        context: dict = {**current_student}

        tpl = DocxTemplate(str(work_template))
        tpl.render(context, autoescape=False)

        output_path = config.OUTPUT_DIR / output_filename
        tpl.save(str(output_path))

    return output_path


def generate_documents_for_all_rows(
    template_path: Path,
    table_name: str,
    filename_column: Optional[str] = None,
    db: Optional[Database] = None,
) -> list[Path]:
    """Генерирует по одному документу на каждую строку данных в таблице."""
    close_db = db is None
    if db is None:
        db = Database()

    try:
        headers, all_rows = db.get_all_data(table_name)
        original_map = db.get_original_headers(table_name)
        all_rows_data = (headers, all_rows)

        generated: list[Path] = []

        for idx, row in enumerate(all_rows):
            row_data = dict(zip(headers, row))

            if filename_column and filename_column in row_data:
                raw_name = str(row_data[filename_column])
                safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", raw_name).strip()
                output_name = f"{safe_name or f'document_{idx + 1:03d}'}.docx"
            else:
                output_name = f"document_{idx + 1:03d}.docx"

            path = generate_document(
                template_path=template_path,
                row_data=row_data,
                output_filename=output_name,
                all_rows_data=all_rows_data,
                original_headers_map=original_map,
            )
            generated.append(path)

        return generated

    finally:
        if close_db:
            db.close()

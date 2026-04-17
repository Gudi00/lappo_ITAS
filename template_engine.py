"""
Движок обработки шаблонов .docx на основе docxtpl (Jinja2).

═══════════════════════════════════════════════════════════════
СИНТАКСИС ШАБЛОНА
═══════════════════════════════════════════════════════════════

1. ТЕКСТОВЫЕ ПОЛЯ (заменяются с сохранением шрифта и форматирования):
   {{ ФИО }}              — значение поля «ФИО» текущего студента
   {{ Группа }}           — значение поля «Группа» текущего студента

2. ТАБЛИЦА ВСЕХ СТУДЕНТОВ (существующая таблица в документе):
   Создай таблицу в Word с заголовком и строкой данных.
   В строке данных используй теги цикла:

   ┌─────────────────────────────────┬───────────┐
   │ ФИО                             │ Группа    │  ← заголовок (форматируй как хочешь)
   ├─────────────────────────────────┼───────────┤
   │ {%tr for s in students %}       │           │  ← первая ячейка строки-цикла
   │ {{ s.ФИО }}                     │ {{ s.Группа }} │
   ├─────────────────────────────────┼───────────┤
   │ {%tr endfor %}                  │           │  ← строка-маркер конца (будет удалена)
   └─────────────────────────────────┴───────────┘

   Проще: весь цикл в одной строке:
   | {%tr for s in students %}{{ s.ФИО }} | {{ s.Группа }}{%tr endfor %} |

3. ТАБЛИЦА ТОЛЬКО ТЕКУЩЕГО СТУДЕНТА:
   Используй current_row вместо students:
   | {%tr for s in current_row %}{{ s.ФИО }} | {{ s.Группа }}{%tr endfor %} |

═══════════════════════════════════════════════════════════════
ПРИМЕР ШАБЛОНА ДИПЛОМА
═══════════════════════════════════════════════════════════════

   Настоящее удостоверение выдано {{ Фамилия Имя Отчество }},
   студенту группы {{ Группа }}.

   Сведения о студенте:
   [таблица с {%tr for s in current_row %}...{%tr endfor %}]

   Список всей группы:
   [таблица с {%tr for s in students %}...{%tr endfor %}]

═══════════════════════════════════════════════════════════════
ПЕРЕМЕННЫЕ КОНТЕКСТА
═══════════════════════════════════════════════════════════════
   {{ ИмяСтолбца }}   — прямой доступ к полю текущего студента
   {{ students }}     — список всех студентов (для цикла по таблице)
   {{ current_row }}  — список из одного текущего студента
"""

import re
from pathlib import Path
from typing import Optional

from docx import Document
from docxtpl import DocxTemplate

import config
from database import Database

# Паттерн для поиска Jinja2-переменных: {{ имя }}
# Игнорирует управляющие теги {% ... %} и точечный доступ {{ s.поле }}
_VAR_PATTERN = re.compile(r"\{\{\s*([^%{}\s][^{}]*?)\s*\}\}")


def find_placeholders(docx_path: Path) -> list[str]:
    """
    Находит все плейсхолдеры {{ ... }} в документе .docx.

    Пропускает управляющие теги ({% for %}, {% if %} и т.д.)
    и переменные с точечным доступом ({{ s.ФИО }} → поле s.ФИО).
    Возвращает отсортированный список уникальных имён.
    """
    doc = Document(str(docx_path))
    placeholders: set[str] = set()

    def _scan(text: str) -> None:
        for m in _VAR_PATTERN.finditer(text):
            name = m.group(1).strip()
            # Пропускаем управляющие конструкции
            if name.startswith(("%", "for ", "if ", "end")):
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
    """
    Строит словарь полей студента.
    Включает как оригинальные (русские) имена, так и санированные (SQLite).
    """
    student: dict[str, str] = {}
    if original_headers_map:
        for orig, safe in original_headers_map.items():
            value = str(row_data.get(safe, ""))
            student[orig] = value
            student[safe] = value
    else:
        student = {k: str(v) for k, v in row_data.items()}
    return student


def generate_document(
    template_path: Path,
    row_data: dict[str, str],
    output_filename: str,
    all_rows_data: Optional[tuple[list[str], list[list]]] = None,
    original_headers_map: Optional[dict[str, str]] = None,
) -> Path:
    """
    Генерирует один документ из шаблона .docx через docxtpl.

    Контекст Jinja2:
      {{ ИмяСтолбца }}   — поле текущего студента (прямой доступ)
      {{ students }}     — список всех студентов для цикла по таблице
      {{ current_row }}  — список из одного текущего студента
    """
    tpl = DocxTemplate(str(template_path))

    # Словарь текущего студента
    current_student = _build_student_dict(row_data, original_headers_map)

    # Список всех студентов для {{ students }}
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
                s[orig] = val_str    # доступ по оригинальному имени: {{ s.ФИО }}
                s[h_raw] = val_str   # доступ по санированному имени: {{ s.fio }}
            students_list.append(s)

    context: dict = {
        **current_student,              # прямой доступ: {{ ФИО }}
        "students": students_list,      # все студенты: {% for s in students %}
        "current_row": [current_student],  # текущий как список: {% for s in current_row %}
    }

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
    """
    Генерирует по одному документу на каждую строку данных в таблице.

    Параметры:
        template_path:    Путь к шаблону .docx
        table_name:       Имя таблицы в SQLite
        filename_column:  Столбец для имени выходного файла (None → «document_001.docx»)
        db:               Экземпляр Database (создаётся автоматически если None)

    Возвращает список путей к сгенерированным файлам.
    """
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

            # Определяем имя выходного файла
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

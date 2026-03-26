"""
Графический интерфейс приложения (tkinter).

Вкладки:
1. Google Sheets — загрузка данных из таблицы
2. Шаблоны — скачивание и заполнение шаблонов
3. Почта — отправка писем
4. База данных — просмотр загруженных данных
"""

import json
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from pathlib import Path
from typing import Optional

import config
from database import Database
from google_sheets import sync_sheet_to_database, extract_sheet_id
from google_docs import download_template, extract_doc_id
from template_engine import (
    find_placeholders,
    generate_document,
    generate_documents_for_all_rows,
)
from email_sender import EmailSender


class App(tk.Tk):
    """Главное окно приложения."""

    def __init__(self):
        super().__init__()

        self.title("Автоматизация документов")
        self.geometry("900x650")
        self.minsize(750, 550)

        # Инициализируем базу данных
        self.db = Database()

        # Создаём вкладки
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self._create_sheets_tab()
        self._create_templates_tab()
        self._create_email_tab()
        self._create_database_tab()

        # Строка состояния
        self.status_var = tk.StringVar(value="Готово")
        status_bar = ttk.Label(
            self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W
        )
        status_bar.pack(fill=tk.X, padx=10, pady=(0, 10))

        # Обработка закрытия окна
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _set_status(self, text: str):
        """Обновляет строку состояния."""
        self.status_var.set(text)
        self.update_idletasks()

    def _run_in_thread(self, target, *args):
        """Запускает функцию в отдельном потоке, чтобы не блокировать GUI."""
        thread = threading.Thread(target=target, args=args, daemon=True)
        thread.start()

    # ================================================================
    # ВКЛАДКА 1: Google Sheets
    # ================================================================
    def _create_sheets_tab(self):
        """Вкладка загрузки данных из Google Sheets."""
        frame = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(frame, text="  Google Sheets  ")

        # URL ввод
        ttk.Label(frame, text="Ссылка на Google Sheets:").pack(anchor=tk.W)
        self.sheet_url_var = tk.StringVar()
        url_entry = ttk.Entry(frame, textvariable=self.sheet_url_var, width=80)
        url_entry.pack(fill=tk.X, pady=(5, 10))

        # Диапазон (опционально)
        range_frame = ttk.Frame(frame)
        range_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(range_frame, text="Диапазон (необязательно):").pack(
            side=tk.LEFT
        )
        self.sheet_range_var = tk.StringVar()
        ttk.Entry(
            range_frame, textvariable=self.sheet_range_var, width=30
        ).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Label(
            range_frame, text="Например: Лист1!A:Z", foreground="gray"
        ).pack(side=tk.LEFT, padx=(10, 0))

        # Кнопка загрузки
        self.btn_load_sheet = ttk.Button(
            frame, text="Загрузить данные", command=self._load_sheet
        )
        self.btn_load_sheet.pack(pady=10)

        # Лог операций
        ttk.Label(frame, text="Журнал:").pack(anchor=tk.W)
        self.sheets_log = scrolledtext.ScrolledText(
            frame, height=12, state=tk.DISABLED, wrap=tk.WORD
        )
        self.sheets_log.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        # Список загруженных таблиц
        ttk.Label(frame, text="Загруженные таблицы:").pack(
            anchor=tk.W, pady=(10, 0)
        )
        self.sheets_listbox = tk.Listbox(frame, height=4)
        self.sheets_listbox.pack(fill=tk.X, pady=(5, 0))
        self._refresh_sheets_list()

    def _log_sheets(self, message: str):
        """Добавляет сообщение в лог вкладки Sheets."""
        self.sheets_log.config(state=tk.NORMAL)
        self.sheets_log.insert(tk.END, message + "\n")
        self.sheets_log.see(tk.END)
        self.sheets_log.config(state=tk.DISABLED)

    def _load_sheet(self):
        """Загружает данные из Google Sheets."""
        url = self.sheet_url_var.get().strip()
        if not url:
            messagebox.showwarning("Внимание", "Вставьте ссылку на таблицу.")
            return

        sheet_id = extract_sheet_id(url)
        if not sheet_id:
            messagebox.showerror("Ошибка", "Некорректная ссылка на Google Sheets.")
            return

        self.btn_load_sheet.config(state=tk.DISABLED)
        self._set_status("Загрузка данных из Google Sheets...")
        self._log_sheets(f"Начинаю загрузку: {url}")

        def do_load():
            try:
                range_name = self.sheet_range_var.get().strip()
                table_name, row_count = sync_sheet_to_database(
                    url, range_name, self.db
                )
                self.after(0, lambda: self._on_sheet_loaded(
                    table_name, row_count
                ))
            except Exception as e:
                self.after(0, lambda msg=str(e): self._on_sheet_error(msg))

        self._run_in_thread(do_load)

    def _on_sheet_loaded(self, table_name: str, row_count: int):
        """Вызывается после успешной загрузки данных."""
        self.btn_load_sheet.config(state=tk.NORMAL)
        self._set_status("Готово")
        self._log_sheets(
            f"Загружено: таблица '{table_name}', строк: {row_count}"
        )
        self._refresh_sheets_list()
        messagebox.showinfo(
            "Успех",
            f"Данные загружены!\n"
            f"Таблица: {table_name}\n"
            f"Строк: {row_count}"
        )

    def _on_sheet_error(self, error: str):
        """Вызывается при ошибке загрузки."""
        self.btn_load_sheet.config(state=tk.NORMAL)
        self._set_status("Ошибка загрузки")
        self._log_sheets(f"ОШИБКА: {error}")
        messagebox.showerror("Ошибка загрузки", error)

    def _refresh_sheets_list(self):
        """Обновляет список загруженных таблиц."""
        self.sheets_listbox.delete(0, tk.END)
        for sheet in self.db.get_all_sheets():
            display = (
                f"{sheet['sheet_name'] or 'Без названия'} "
                f"({sheet['table_name']}, "
                f"строк: {sheet['row_count']})"
            )
            self.sheets_listbox.insert(tk.END, display)

    # ================================================================
    # ВКЛАДКА 2: Шаблоны
    # ================================================================
    def _create_templates_tab(self):
        """Вкладка работы с шаблонами Google Docs."""
        frame = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(frame, text="  Шаблоны  ")

        # URL шаблона
        ttk.Label(frame, text="Ссылка на шаблон Google Docs:").pack(
            anchor=tk.W
        )
        self.doc_url_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.doc_url_var, width=80).pack(
            fill=tk.X, pady=(5, 10)
        )

        # Кнопка скачивания шаблона
        self.btn_download_template = ttk.Button(
            frame, text="Скачать шаблон",
            command=self._download_template
        )
        self.btn_download_template.pack(pady=5)

        # Выбор таблицы для данных
        select_frame = ttk.Frame(frame)
        select_frame.pack(fill=tk.X, pady=10)

        ttk.Label(select_frame, text="Таблица с данными:").pack(
            side=tk.LEFT
        )
        self.template_table_var = tk.StringVar()
        self.template_table_combo = ttk.Combobox(
            select_frame, textvariable=self.template_table_var,
            state="readonly", width=40
        )
        self.template_table_combo.pack(side=tk.LEFT, padx=(10, 0))

        ttk.Button(
            select_frame, text="Обновить", command=self._refresh_table_combo
        ).pack(side=tk.LEFT, padx=(5, 0))

        # Столбец для имени файла
        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(
            name_frame,
            text="Столбец для имени файла (необязательно):"
        ).pack(side=tk.LEFT)
        self.filename_column_var = tk.StringVar()
        ttk.Entry(
            name_frame, textvariable=self.filename_column_var, width=25
        ).pack(side=tk.LEFT, padx=(10, 0))


        # Кнопка генерации
        self.btn_generate = ttk.Button(
            frame, text="Сгенерировать документы",
            command=self._generate_documents
        )
        self.btn_generate.pack(pady=5)

        # Лог
        ttk.Label(frame, text="Журнал:").pack(anchor=tk.W, pady=(10, 0))
        self.templates_log = scrolledtext.ScrolledText(
            frame, height=10, state=tk.DISABLED, wrap=tk.WORD
        )
        self.templates_log.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        self._refresh_table_combo()

    def _log_templates(self, message: str):
        """Добавляет сообщение в лог вкладки Шаблоны."""
        self.templates_log.config(state=tk.NORMAL)
        self.templates_log.insert(tk.END, message + "\n")
        self.templates_log.see(tk.END)
        self.templates_log.config(state=tk.DISABLED)

    def _refresh_table_combo(self):
        """Обновляет выпадающий список таблиц."""
        tables = self.db.get_table_names()
        self.template_table_combo["values"] = tables
        if tables and not self.template_table_var.get():
            self.template_table_var.set(tables[0])

    def _download_template(self):
        """Скачивает шаблон из Google Docs."""
        url = self.doc_url_var.get().strip()
        if not url:
            messagebox.showwarning(
                "Внимание", "Вставьте ссылку на шаблон Google Docs."
            )
            return

        if not extract_doc_id(url):
            messagebox.showerror(
                "Ошибка", "Некорректная ссылка на Google Docs."
            )
            return

        self.btn_download_template.config(state=tk.DISABLED)
        self._set_status("Скачивание шаблона...")
        self._log_templates(f"Скачиваю шаблон: {url}")

        def do_download():
            try:
                local_path = download_template(url, self.db)
                placeholders = find_placeholders(local_path)
                self.after(0, lambda: self._on_template_downloaded(
                    local_path, placeholders
                ))
            except Exception as e:
                self.after(0, lambda msg=str(e): self._on_template_error(msg))

        self._run_in_thread(do_download)

    def _on_template_downloaded(self, path: Path, placeholders: list[str]):
        """Вызывается после успешного скачивания шаблона."""
        self.btn_download_template.config(state=tk.NORMAL)
        self._set_status("Готово")
        self._log_templates(f"Шаблон сохранён: {path}")
        self._log_templates(
            f"Найдены плейсхолдеры: {', '.join(placeholders) or 'нет'}"
        )

    def _on_template_error(self, error: str):
        """Вызывается при ошибке скачивания шаблона."""
        self.btn_download_template.config(state=tk.NORMAL)
        self._set_status("Ошибка")
        self._log_templates(f"ОШИБКА: {error}")
        messagebox.showerror("Ошибка", error)

    def _generate_documents(self):
        """Генерирует документы из шаблона."""
        table_name = self.template_table_var.get()
        if not table_name:
            messagebox.showwarning(
                "Внимание", "Выберите таблицу с данными."
            )
            return

        # Ищем последний скачанный шаблон
        templates = self.db.get_all_templates()
        if not templates:
            messagebox.showwarning(
                "Внимание", "Сначала скачайте шаблон."
            )
            return

        template_info = templates[0]  # Последний скачанный
        template_path = Path(template_info["local_path"])

        if not template_path.exists():
            messagebox.showerror(
                "Ошибка",
                f"Файл шаблона не найден: {template_path}\n"
                f"Скачайте шаблон заново."
            )
            return

        self.btn_generate.config(state=tk.DISABLED)
        self._set_status("Генерация документов...")

        def do_generate():
            try:
                filename_col = self.filename_column_var.get().strip() or None
                files = generate_documents_for_all_rows(
                    template_path=template_path,
                    table_name=table_name,
                    filename_column=filename_col,
                    db=self.db
                )
                self.after(0, lambda: self._on_documents_generated(files))
            except Exception as e:
                self.after(0, lambda msg=str(e): self._on_generate_error(msg))

        self._run_in_thread(do_generate)

    def _on_documents_generated(self, files: list[Path]):
        """Вызывается после успешной генерации документов."""
        self.btn_generate.config(state=tk.NORMAL)
        self._set_status("Готово")
        self._log_templates(f"Сгенерировано документов: {len(files)}")
        for f in files:
            self._log_templates(f"  -> {f}")
        messagebox.showinfo(
            "Успех",
            f"Сгенерировано {len(files)} документов.\n"
            f"Папка: {config.OUTPUT_DIR}"
        )

    def _on_generate_error(self, error: str):
        """Вызывается при ошибке генерации."""
        self.btn_generate.config(state=tk.NORMAL)
        self._set_status("Ошибка генерации")
        self._log_templates(f"ОШИБКА: {error}")
        messagebox.showerror("Ошибка генерации", error)

    # ================================================================
    # ВКЛАДКА 3: Почта
    # ================================================================
    def _create_email_tab(self):
        """Вкладка отправки писем."""
        frame = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(frame, text="  Почта  ")

        # Получатель
        ttk.Label(frame, text="Получатель (или несколько через ;):").pack(
            anchor=tk.W
        )
        self.email_to_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.email_to_var, width=60).pack(
            fill=tk.X, pady=(5, 10)
        )

        # Тема
        ttk.Label(frame, text="Тема письма:").pack(anchor=tk.W)
        self.email_subject_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.email_subject_var, width=60).pack(
            fill=tk.X, pady=(5, 10)
        )

        # Текст письма
        ttk.Label(frame, text="Текст письма:").pack(anchor=tk.W)
        self.email_body_text = scrolledtext.ScrolledText(
            frame, height=8, wrap=tk.WORD
        )
        self.email_body_text.pack(fill=tk.BOTH, expand=True, pady=(5, 10))

        # Прикрепить файл
        attach_frame = ttk.Frame(frame)
        attach_frame.pack(fill=tk.X, pady=(0, 10))
        self.attachment_var = tk.StringVar()
        ttk.Label(attach_frame, text="Вложение:").pack(side=tk.LEFT)
        ttk.Entry(
            attach_frame, textvariable=self.attachment_var, width=50
        ).pack(side=tk.LEFT, padx=(10, 5))
        ttk.Button(
            attach_frame, text="Выбрать...",
            command=self._select_attachment
        ).pack(side=tk.LEFT)

        # Галочка тестового режима
        self.dry_run_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame,
            text="Тестовый режим (письма НЕ отправляются)",
            variable=self.dry_run_var
        ).pack(anchor=tk.W, pady=(0, 10))

        # Кнопки
        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        self.btn_send_email = ttk.Button(
            btn_frame, text="Отправить", command=self._send_email
        )
        self.btn_send_email.pack(side=tk.LEFT, padx=5)

        # Лог отправки
        ttk.Label(frame, text="Журнал отправки:").pack(
            anchor=tk.W, pady=(10, 0)
        )
        self.email_log_text = scrolledtext.ScrolledText(
            frame, height=5, state=tk.DISABLED, wrap=tk.WORD
        )
        self.email_log_text.pack(fill=tk.X, pady=(5, 0))

    def _log_email(self, message: str):
        """Добавляет сообщение в лог вкладки Почта."""
        self.email_log_text.config(state=tk.NORMAL)
        self.email_log_text.insert(tk.END, message + "\n")
        self.email_log_text.see(tk.END)
        self.email_log_text.config(state=tk.DISABLED)

    def _select_attachment(self):
        """Открывает диалог выбора файла для вложения."""
        path = filedialog.askopenfilename(
            title="Выберите файл для вложения",
            filetypes=[
                ("Документы Word", "*.docx"),
                ("PDF файлы", "*.pdf"),
                ("Все файлы", "*.*"),
            ]
        )
        if path:
            self.attachment_var.set(path)

    def _send_email(self):
        """Отправляет письмо."""
        recipients_str = self.email_to_var.get().strip()
        subject = self.email_subject_var.get().strip()
        body = self.email_body_text.get("1.0", tk.END).strip()

        if not recipients_str:
            messagebox.showwarning("Внимание", "Укажите получателя.")
            return
        if not subject:
            messagebox.showwarning("Внимание", "Укажите тему письма.")
            return
        if not body:
            messagebox.showwarning("Внимание", "Введите текст письма.")
            return

        self.btn_send_email.config(state=tk.DISABLED)
        self._set_status("Отправка письма...")

        def do_send():
            try:
                sender = EmailSender(
                    db=self.db, dry_run=self.dry_run_var.get()
                )
                attachment = self.attachment_var.get().strip() or None
                attachment_path = Path(attachment) if attachment else None

                # Разделяем получателей по ;
                recipients = [
                    r.strip() for r in recipients_str.split(";") if r.strip()
                ]

                results = {"sent": 0, "failed": 0, "errors": []}
                for recipient in recipients:
                    try:
                        sender.send_email(
                            to=recipient,
                            subject=subject,
                            body=body,
                            attachment_path=attachment_path
                        )
                        results["sent"] += 1
                    except Exception as e:
                        results["failed"] += 1
                        results["errors"].append(f"{recipient}: {str(e)}")

                self.after(0, lambda: self._on_email_sent(results))
            except Exception as e:
                self.after(0, lambda msg=str(e): self._on_email_error(msg))

        self._run_in_thread(do_send)

    def _on_email_sent(self, results: dict):
        """Вызывается после попытки отправки."""
        self.btn_send_email.config(state=tk.NORMAL)
        self._set_status("Готово")

        mode = "ТЕСТ" if self.dry_run_var.get() else "Отправлено"
        self._log_email(
            f"[{mode}] Успешно: {results['sent']}, "
            f"Ошибок: {results['failed']}"
        )
        for err in results["errors"]:
            self._log_email(f"  Ошибка: {err}")

        if results["failed"] == 0:
            messagebox.showinfo(
                "Результат",
                f"{'Тест' if self.dry_run_var.get() else 'Отправка'} "
                f"завершена.\nУспешно: {results['sent']}"
            )
        else:
            messagebox.showwarning(
                "Результат",
                f"Успешно: {results['sent']}\n"
                f"Ошибок: {results['failed']}"
            )

    def _on_email_error(self, error: str):
        """Вызывается при ошибке отправки."""
        self.btn_send_email.config(state=tk.NORMAL)
        self._set_status("Ошибка отправки")
        self._log_email(f"ОШИБКА: {error}")
        messagebox.showerror("Ошибка", error)

    # ================================================================
    # ВКЛАДКА 4: База данных
    # ================================================================
    def _create_database_tab(self):
        """Вкладка просмотра данных из БД."""
        frame = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(frame, text="  База данных  ")

        # Выбор таблицы
        select_frame = ttk.Frame(frame)
        select_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(select_frame, text="Таблица:").pack(side=tk.LEFT)
        self.db_table_var = tk.StringVar()
        self.db_table_combo = ttk.Combobox(
            select_frame, textvariable=self.db_table_var,
            state="readonly", width=40
        )
        self.db_table_combo.pack(side=tk.LEFT, padx=(10, 5))

        ttk.Button(
            select_frame, text="Обновить список",
            command=self._refresh_db_tables
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            select_frame, text="Показать данные",
            command=self._show_table_data
        ).pack(side=tk.LEFT, padx=5)

        # Таблица данных (Treeview)
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.data_tree = ttk.Treeview(tree_frame, show="headings")
        self.data_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Полосы прокрутки
        v_scroll = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.data_tree.yview
        )
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.data_tree.configure(yscrollcommand=v_scroll.set)

        h_scroll = ttk.Scrollbar(
            frame, orient=tk.HORIZONTAL, command=self.data_tree.xview
        )
        h_scroll.pack(fill=tk.X)
        self.data_tree.configure(xscrollcommand=h_scroll.set)

        # Информация о таблице
        self.db_info_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.db_info_var).pack(
            anchor=tk.W, pady=(5, 0)
        )

        self._refresh_db_tables()

    def _refresh_db_tables(self):
        """Обновляет список таблиц в выпадающем списке."""
        tables = self.db.get_table_names()
        self.db_table_combo["values"] = tables
        # Также обновляем комбобокс на вкладке шаблонов
        self.template_table_combo["values"] = tables
        if tables and not self.db_table_var.get():
            self.db_table_var.set(tables[0])

    def _show_table_data(self):
        """Отображает данные из выбранной таблицы."""
        table_name = self.db_table_var.get()
        if not table_name:
            messagebox.showwarning("Внимание", "Выберите таблицу.")
            return

        try:
            headers, rows = self.db.get_all_data(table_name)
            original_map = self.db.get_original_headers(table_name)

            # Получаем оригинальные имена для отображения
            reverse_map = {v: k for k, v in original_map.items()}

            # Очищаем таблицу
            self.data_tree.delete(*self.data_tree.get_children())
            self.data_tree["columns"] = headers

            for col in headers:
                display_name = reverse_map.get(col, col)
                self.data_tree.heading(col, text=display_name)
                self.data_tree.column(col, width=120, minwidth=60)

            for row in rows:
                self.data_tree.insert("", tk.END, values=row)

            self.db_info_var.set(
                f"Таблица: {table_name} | "
                f"Столбцов: {len(headers)} | "
                f"Строк: {len(rows)}"
            )

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные:\n{e}")

    # ================================================================
    # Общие методы
    # ================================================================
    def _on_close(self):
        """Обработка закрытия приложения."""
        self.db.close()
        self.destroy()

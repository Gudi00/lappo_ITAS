"""
Графический интерфейс приложения в стиле macOS.
Использует CustomTkinter для современного внешнего вида.
"""

import re
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, filedialog
from tkinter import ttk as tkttk
from pathlib import Path

import customtkinter as ctk

import config
from database import Database
from google_sheets import sync_sheet_to_database, extract_sheet_id
from google_docs import download_template, extract_doc_id
from template_engine import find_placeholders, generate_documents_for_all_rows, generate_document
from email_sender import EmailSender

# ── Тема ──────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Цветовая палитра macOS
_BG       = "#F2F2F7"   # фон окна / контент
_SIDEBAR  = "#FFFFFF"   # боковая панель
_CARD     = "#FFFFFF"   # карточки
_ACCENT   = "#007AFF"   # синий Apple
_ACCENT_H = "#0062CC"   # hover синий
_TEXT1    = "#1C1C1E"   # основной текст
_TEXT2    = "#8E8E93"   # второстепенный текст
_BORDER   = "#D1D1D6"   # граница
_SUCCESS  = "#34C759"   # зелёный
_ERROR    = "#FF3B30"   # красный
_ODD_ROW  = "#F9F9FB"   # нечётная строка таблицы

_FONT     = "Helvetica Neue"
_MONO     = "Menlo"


# ── Вспомогательные виджеты ───────────────────────────────────────────────────

class _NavButton(ctk.CTkButton):
    """Кнопка бокового меню."""

    def __init__(self, master, label: str, icon: str, **kw):
        super().__init__(
            master,
            text=f"  {icon}   {label}",
            anchor="w",
            height=40,
            corner_radius=8,
            fg_color="transparent",
            text_color=_TEXT1,
            hover_color="#EBEBF0",
            font=ctk.CTkFont(family=_FONT, size=14),
            **kw,
        )

    def activate(self):
        self.configure(fg_color=_ACCENT, text_color="white", hover_color=_ACCENT_H)

    def deactivate(self):
        self.configure(fg_color="transparent", text_color=_TEXT1, hover_color="#EBEBF0")


class _TablePickerWidget(ctk.CTkFrame):
    """
    Compact table-selector: clickable rows with hover-preview popup.
    The preview appears after a short delay and shows the first rows of data.
    selected_var always holds the *internal* SQLite table name.
    """
    _DELAY_MS = 380

    def __init__(self, parent, db, selected_var: tk.StringVar, **kw):
        super().__init__(
            parent, fg_color="white", corner_radius=8,
            border_width=1, border_color=_BORDER, **kw,
        )
        self._db = db
        self._sel = selected_var
        self._tables: list[tuple[str, str]] = []
        self._popup = None
        self._after_id = None      # timer: show popup
        self._hide_after_id = None # timer: hide popup

        self._inner = tk.Frame(self, bg="white")
        self._inner.pack(fill="both", expand=True, padx=1, pady=1)

    # ── public ────────────────────────────────────────────────────────────────

    def refresh(self, tables: list[tuple[str, str]]) -> None:
        """Re-populate list with [(internal_name, display_name), ...]."""
        self._tables = tables
        internals = [t[0] for t in tables]
        if self._sel.get() not in internals and internals:
            self._sel.set(internals[0])
        self._rebuild()

    # ── internal ──────────────────────────────────────────────────────────────

    def _rebuild(self) -> None:
        for w in self._inner.winfo_children():
            w.destroy()
        if not self._tables:
            tk.Label(
                self._inner, text="Нет загруженных таблиц",
                bg="white", fg=_TEXT2, font=(_FONT, 12), pady=10,
            ).pack()
            return
        sel = self._sel.get()
        for i, (internal, display) in enumerate(self._tables):
            is_sel = internal == sel
            is_last = i == len(self._tables) - 1
            self._add_row(internal, display, is_sel, is_last)

    def _add_row(self, internal: str, display: str, selected: bool, last: bool) -> None:
        bg    = _ACCENT if selected else "white"
        fg1   = "white" if selected else _TEXT1
        fg2   = "white" if selected else _TEXT2

        row = tk.Frame(self._inner, bg=bg, cursor="hand2")
        row.pack(fill="x")

        tk.Label(row, text="⊞", bg=bg, fg=fg1,
                 font=(_FONT, 12), padx=10, pady=8).pack(side="left")
        tk.Label(row, text=display, bg=bg, fg=fg1,
                 font=(_FONT, 13), anchor="w", pady=8).pack(
                     side="left", fill="x", expand=True)
        try:
            cnt = self._db.get_row_count(internal)
            detail = f"{cnt} стр."
        except Exception:
            detail = ""
        tk.Label(row, text=detail, bg=bg, fg=fg2,
                 font=(_FONT, 11), padx=12, pady=8).pack(side="right")

        if not last:
            tk.Frame(self._inner, bg=_BORDER, height=1).pack(fill="x")

        for w in (*row.winfo_children(), row):
            w.bind("<Button-1>", lambda e, n=internal: self._on_click(n))
            w.bind("<Enter>",    lambda e, n=internal, r=row, s=selected:
                                 self._on_enter(e, n, r, s))
            w.bind("<Leave>",    lambda e, r=row, s=selected:
                                 self._on_leave(r, s))

    def _on_click(self, internal: str) -> None:
        self._hide_popup()
        self._sel.set(internal)
        self._rebuild()

    def _on_enter(self, event, internal: str, row: tk.Frame, is_sel: bool) -> None:
        if not is_sel:
            self._tint(row, "#EBF4FF", _TEXT1, _TEXT2)
        if self._after_id:
            self.after_cancel(self._after_id)
        self._after_id = self.after(
            self._DELAY_MS,
            lambda: self._show_popup(internal, event.x_root, event.y_root),
        )

    def _on_leave(self, row: tk.Frame, is_sel: bool) -> None:
        if not is_sel:
            self._tint(row, "white", _TEXT1, _TEXT2)
        if self._after_id:
            self.after_cancel(self._after_id)
            self._after_id = None
        if self._hide_after_id:
            self.after_cancel(self._hide_after_id)
        self._hide_after_id = self.after(600, self._maybe_hide)

    def _tint(self, row: tk.Frame, bg: str, fg: str, fg2: str) -> None:
        try:
            row.configure(bg=bg)
            ch = row.winfo_children()
            for i, w in enumerate(ch):
                w.configure(bg=bg, fg=(fg2 if i == len(ch) - 1 else fg))
        except Exception:
            pass

    def _cancel_hide(self) -> None:
        """Cancel a pending hide timer (called when cursor enters the popup)."""
        if self._hide_after_id:
            self.after_cancel(self._hide_after_id)
            self._hide_after_id = None

    def _maybe_hide(self) -> None:
        self._hide_after_id = None
        if self._popup is None:
            return
        try:
            if not self._popup.winfo_exists():
                self._popup = None
                return
            px, py = self._popup.winfo_rootx(), self._popup.winfo_rooty()
            pw = max(self._popup.winfo_width(), self._popup.winfo_reqwidth())
            ph = max(self._popup.winfo_height(), self._popup.winfo_reqheight())
            mx = self._popup.winfo_pointerx()
            my = self._popup.winfo_pointery()
            if px <= mx <= px + pw and py <= my <= py + ph:
                return
        except Exception:
            pass
        self._hide_popup()

    def _hide_popup(self) -> None:
        if self._after_id:
            self.after_cancel(self._after_id)
            self._after_id = None
        if self._hide_after_id:
            self.after_cancel(self._hide_after_id)
            self._hide_after_id = None
        if self._popup is not None:
            try:
                if self._popup.winfo_exists():
                    self._popup.destroy()
            except Exception:
                pass
        self._popup = None

    def _show_popup(self, internal: str, rx: int, ry: int) -> None:
        self._hide_popup()
        try:
            headers, rows = self._db.get_all_data(internal)
            orig    = self._db.get_original_headers(internal)
            aliases = self._db.get_column_aliases(internal)
            rev     = {v: k for k, v in orig.items()}
        except Exception:
            return
        if not headers:
            return

        cols      = headers[:5]
        prev_rows = rows[:6]

        popup = tk.Toplevel()
        popup.overrideredirect(True)

        outer = tk.Frame(popup, bg=_BORDER, padx=1, pady=1)
        outer.pack(fill="both", expand=True)
        inner = tk.Frame(outer, bg="white")
        inner.pack(fill="both", expand=True)

        # Header strip
        try:
            all_tbls  = self._db.get_tables_with_display_names()
            disp_name = next((d for t, d in all_tbls if t == internal), internal)
        except Exception:
            disp_name = internal
        try:
            total = self._db.get_row_count(internal)
        except Exception:
            total = len(rows)

        hdr = tk.Frame(inner, bg=_BG)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"  {disp_name}", bg=_BG, fg=_TEXT1,
                 font=(_FONT, 12, "bold"), anchor="w", pady=7).pack(side="left")
        tk.Label(hdr, text=f"{total} строк  ", bg=_BG, fg=_TEXT2,
                 font=(_FONT, 11), anchor="e", pady=7).pack(side="right")
        tk.Frame(inner, bg=_BORDER, height=1).pack(fill="x")

        # Mini Treeview
        _SN = "PopPrev.Treeview"
        sty = tkttk.Style()
        sty.configure(_SN, rowheight=22, font=(_FONT, 11),
                      background="white", fieldbackground="white",
                      foreground=_TEXT1, borderwidth=0)
        sty.configure(f"{_SN}.Heading", font=(_FONT, 10, "bold"),
                      background=_BG, foreground=_TEXT2, relief="flat")
        sty.map(_SN,
                background=[("selected", "#E3EEFB")],
                foreground=[("selected", _TEXT1)])

        col_widths = []
        for col in cols:
            disp = aliases.get(col) or rev.get(col, col)
            max_vl = max(
                (len(str(r[headers.index(col)])) for r in prev_rows
                 if headers.index(col) < len(r)),
                default=0,
            )
            col_widths.append(min(max(len(disp) * 8 + 20, max_vl * 7 + 10, 70), 160))

        tree = tkttk.Treeview(inner, columns=cols, show="headings",
                               style=_SN, height=min(len(prev_rows), 6))
        for col, w in zip(cols, col_widths):
            disp = aliases.get(col) or rev.get(col, col)
            tree.heading(col, text=disp)
            tree.column(col, width=w, minwidth=40, stretch=False)

        for row_data in prev_rows:
            vals = [
                str(row_data[headers.index(c)])[:28] if c in headers else ""
                for c in cols
            ]
            tree.insert("", "end", values=vals)
        tree.pack(padx=4, pady=4)

        if len(rows) > 6:
            more = total - 6 if isinstance(total, int) else "…"
            tk.Label(inner, text=f"  … ещё {more} строк",
                     bg="white", fg=_TEXT2, font=(_FONT, 10),
                     anchor="w", pady=4).pack(fill="x", padx=8)

        popup.update_idletasks()
        pw  = popup.winfo_reqwidth()
        ph  = popup.winfo_reqheight()
        sw  = popup.winfo_screenwidth()
        sh  = popup.winfo_screenheight()
        px  = rx + 16
        py  = ry - 30
        if px + pw > sw - 10:
            px = rx - pw - 16
        if py + ph > sh - 10:
            py = sh - ph - 10
        popup.geometry(f"+{max(0, px)}+{max(0, py)}")
        self._popup = popup
        popup.bind("<Enter>", lambda e: self._cancel_hide())
        popup.bind("<Leave>", lambda e: self.after(300, self._maybe_hide))


class _ColumnRenameDialog(ctk.CTkToplevel):
    """Диалог массового переименования столбцов."""

    def __init__(self, parent, db, table_name: str, on_save):
        super().__init__(parent)
        self.title("Переименование столбцов")
        self.geometry("580x500")
        self.resizable(False, True)
        self.configure(fg_color=_BG)
        self.after(150, lambda: self.grab_set() if self.winfo_exists() else None)

        self._db = db
        self._table_name = table_name
        self._on_save = on_save
        self._entries: list[tuple[str, tk.StringVar]] = []
        self._build()

    def _build(self):
        ctk.CTkLabel(
            self, text="Переименование столбцов",
            font=ctk.CTkFont(family=_FONT, size=18, weight="bold"),
            text_color=_TEXT1,
        ).pack(anchor="w", padx=24, pady=(20, 2))
        ctk.CTkLabel(
            self,
            text="Локальные имена сохраняются при повторной загрузке таблицы.",
            font=ctk.CTkFont(family=_FONT, size=12),
            text_color=_TEXT2,
        ).pack(anchor="w", padx=24, pady=(0, 10))

        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=24)
        ctk.CTkLabel(hdr, text="Исходное имя (Google Sheets)", width=240, anchor="w",
                     font=ctk.CTkFont(family=_FONT, size=12, weight="bold"),
                     text_color=_TEXT2).pack(side="left")
        ctk.CTkLabel(hdr, text="Локальное имя", anchor="w",
                     font=ctk.CTkFont(family=_FONT, size=12, weight="bold"),
                     text_color=_TEXT2).pack(side="left", padx=(8, 0))
        ctk.CTkFrame(self, height=1, fg_color=_BORDER).pack(fill="x", padx=24, pady=(6, 0))

        scroll = ctk.CTkScrollableFrame(
            self, fg_color="white", corner_radius=8,
            border_width=1, border_color=_BORDER,
        )
        scroll.pack(fill="both", expand=True, padx=24, pady=(4, 12))

        try:
            headers, _ = self._db.get_all_data(self._table_name)
            orig = self._db.get_original_headers(self._table_name)
            aliases = self._db.get_column_aliases(self._table_name)
            rev = {v: k for k, v in orig.items()}
        except Exception:
            headers, rev, aliases = [], {}, {}

        for safe_col in headers:
            orig_name = rev.get(safe_col, safe_col)
            var = tk.StringVar(value=aliases.get(safe_col, ""))
            self._entries.append((safe_col, var))

            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=4, padx=8)
            ctk.CTkLabel(
                row, text=orig_name, width=240, anchor="w",
                font=ctk.CTkFont(family=_FONT, size=13),
                text_color=_TEXT1,
            ).pack(side="left")
            ctk.CTkEntry(
                row, textvariable=var, width=250,
                corner_radius=6, border_width=1, border_color=_BORDER,
                fg_color="white", text_color=_TEXT1,
                placeholder_text=orig_name,
                font=ctk.CTkFont(family=_FONT, size=13),
            ).pack(side="left", padx=(8, 0))

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=24, pady=(0, 20))
        ctk.CTkButton(
            btn_row, text="Отмена", command=self.destroy,
            corner_radius=8, height=36, width=100,
            fg_color="#E8E8ED", hover_color="#D8D8DD", text_color=_TEXT1,
            font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
        ).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            btn_row, text="Сохранить", command=self._save,
            corner_radius=8, height=36, width=110,
            fg_color=_ACCENT, hover_color=_ACCENT_H, text_color="white",
            font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
        ).pack(side="right")

    def _save(self):
        for safe_col, var in self._entries:
            self._db.set_column_alias(self._table_name, safe_col, var.get())
        self.destroy()
        if self._on_save:
            self._on_save()


class _MergeColumnsDialog(ctk.CTkToplevel):
    """Диалог объединения двух столбцов."""

    def __init__(self, parent, db, table_name: str, on_merge):
        super().__init__(parent)
        self.title("Объединение столбцов")
        self.geometry("480x360")
        self.resizable(False, False)
        self.configure(fg_color=_BG)
        self.after(150, lambda: self.grab_set() if self.winfo_exists() else None)

        self._db = db
        self._table_name = table_name
        self._on_merge = on_merge
        self._col_map: dict[str, str] = {}
        self._build()

    def _build(self):
        ctk.CTkLabel(
            self, text="Объединение столбцов",
            font=ctk.CTkFont(family=_FONT, size=18, weight="bold"),
            text_color=_TEXT1,
        ).pack(anchor="w", padx=24, pady=(20, 2))
        ctk.CTkLabel(
            self,
            text="Оба непустых — объединяются через пробел.\nОдин пустой — берётся значение из другого.",
            font=ctk.CTkFont(family=_FONT, size=12),
            text_color=_TEXT2, justify="left",
        ).pack(anchor="w", padx=24, pady=(0, 16))

        try:
            headers, _ = self._db.get_all_data(self._table_name)
            orig = self._db.get_original_headers(self._table_name)
            aliases = self._db.get_column_aliases(self._table_name)
            rev = {v: k for k, v in orig.items()}
        except Exception:
            headers, rev, aliases = [], {}, {}

        display_names = []
        for safe in headers:
            disp = aliases.get(safe) or rev.get(safe, safe)
            display_names.append(disp)
            self._col_map[disp] = safe

        r1 = ctk.CTkFrame(self, fg_color="transparent")
        r1.pack(fill="x", padx=24, pady=(0, 10))
        ctk.CTkLabel(r1, text="Первый столбец:", width=160, anchor="w",
                     font=ctk.CTkFont(family=_FONT, size=13), text_color=_TEXT1).pack(side="left")
        self._col_a_var = tk.StringVar(value=display_names[0] if display_names else "")
        ctk.CTkComboBox(
            r1, variable=self._col_a_var, values=display_names, state="readonly", width=250,
            corner_radius=8, border_color=_BORDER, fg_color="white",
            button_color=_BORDER, button_hover_color=_TEXT2,
            font=ctk.CTkFont(family=_FONT, size=13),
            dropdown_font=ctk.CTkFont(family=_FONT, size=13),
        ).pack(side="left")

        r2 = ctk.CTkFrame(self, fg_color="transparent")
        r2.pack(fill="x", padx=24, pady=(0, 16))
        ctk.CTkLabel(r2, text="Второй столбец:", width=160, anchor="w",
                     font=ctk.CTkFont(family=_FONT, size=13), text_color=_TEXT1).pack(side="left")
        self._col_b_var = tk.StringVar(value=display_names[1] if len(display_names) > 1 else "")
        ctk.CTkComboBox(
            r2, variable=self._col_b_var, values=display_names, state="readonly", width=250,
            corner_radius=8, border_color=_BORDER, fg_color="white",
            button_color=_BORDER, button_hover_color=_TEXT2,
            font=ctk.CTkFont(family=_FONT, size=13),
            dropdown_font=ctk.CTkFont(family=_FONT, size=13),
        ).pack(side="left")

        warn = ctk.CTkFrame(self, fg_color="#FFF8E1", corner_radius=8,
                            border_width=1, border_color="#FFD54F")
        warn.pack(fill="x", padx=24, pady=(0, 20))
        ctk.CTkLabel(
            warn,
            text="⚠  Результат записывается в первый столбец.\n"
                 "   Второй столбец будет удалён из таблицы.",
            font=ctk.CTkFont(family=_FONT, size=12),
            text_color="#5D4037", justify="left", anchor="w",
        ).pack(padx=12, pady=10, anchor="w")

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=24, pady=(0, 20))
        ctk.CTkButton(
            btn_row, text="Отмена", command=self.destroy,
            corner_radius=8, height=36, width=100,
            fg_color="#E8E8ED", hover_color="#D8D8DD", text_color=_TEXT1,
            font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
        ).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            btn_row, text="Объединить", command=self._merge,
            corner_radius=8, height=36, width=120,
            fg_color=_ACCENT, hover_color=_ACCENT_H, text_color="white",
            font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
        ).pack(side="right")

    def _merge(self):
        disp_a = self._col_a_var.get()
        disp_b = self._col_b_var.get()
        if not disp_a or not disp_b:
            messagebox.showwarning("Внимание", "Выберите оба столбца.", parent=self)
            return
        if disp_a == disp_b:
            messagebox.showwarning("Внимание", "Выберите разные столбцы.", parent=self)
            return
        col_a = self._col_map.get(disp_a, disp_a)
        col_b = self._col_map.get(disp_b, disp_b)
        try:
            self._db.merge_columns(self._table_name, col_a, col_b)
            self.destroy()
            if self._on_merge:
                self._on_merge()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось объединить:\n{e}", parent=self)


# ── Главное окно ──────────────────────────────────────────────────────────────

class App(ctk.CTk):
    """Главное окно приложения."""

    def __init__(self):
        super().__init__()

        self.title("DocFlow")
        self.geometry("1060x700")
        self.minsize(860, 580)
        self.configure(fg_color=_BG)

        self.db = Database()

        self._build_ui()
        self._show("sheets")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Построение UI ─────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Боковая панель ──
        sidebar = ctk.CTkFrame(
            self, width=210, corner_radius=0,
            fg_color=_SIDEBAR,
            border_width=1, border_color=_BORDER,
        )
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # Логотип / название
        ctk.CTkLabel(
            sidebar, text="DocFlow",
            font=ctk.CTkFont(family=_FONT, size=22, weight="bold"),
            text_color=_TEXT1,
        ).pack(padx=20, pady=(28, 2))

        ctk.CTkLabel(
            sidebar, text="Автоматизация документов",
            font=ctk.CTkFont(family=_FONT, size=11),
            text_color=_TEXT2,
        ).pack(padx=20, pady=(0, 20))

        ctk.CTkFrame(sidebar, height=1, fg_color=_BORDER).pack(fill="x", padx=16, pady=(0, 14))

        # Навигация
        nav = [
            ("sheets",    "Google Sheets", "⊞"),
            ("templates", "Шаблоны",       "⎘"),
            ("email",     "Почта",         "✉"),
            ("database",  "База данных",   "≡"),
        ]
        self._nav_btns: dict[str, _NavButton] = {}
        for key, label, icon in nav:
            btn = _NavButton(
                sidebar, label=label, icon=icon,
                command=lambda k=key: self._show(k),
            )
            btn.pack(fill="x", padx=10, pady=2)
            self._nav_btns[key] = btn

        # Статус — внизу боковой панели
        ctk.CTkFrame(sidebar, height=1, fg_color=_BORDER).pack(
            side="bottom", fill="x", padx=16, pady=(0, 12))
        self._status_dot = ctk.CTkLabel(
            sidebar, text="● Готово",
            font=ctk.CTkFont(family=_FONT, size=12),
            text_color=_SUCCESS, anchor="w",
        )
        self._status_dot.pack(side="bottom", padx=20, pady=(0, 16), fill="x")

        # ── Контентная область ──
        self._content = ctk.CTkFrame(self, corner_radius=0, fg_color=_BG)
        self._content.pack(side="left", fill="both", expand=True)

        # Секции
        self._sections: dict[str, ctk.CTkFrame] = {}
        builders = {
            "sheets":    self._build_sheets,
            "templates": self._build_templates,
            "email":     self._build_email,
            "database":  self._build_database,
        }
        for key, build_fn in builders.items():
            f = ctk.CTkScrollableFrame(
                self._content,
                corner_radius=0, fg_color=_BG,
                scrollbar_button_color=_BORDER,
                scrollbar_button_hover_color=_TEXT2,
            )
            build_fn(f)
            self._sections[key] = f

    def _show(self, key: str):
        for k, f in self._sections.items():
            f.pack_forget()
        self._sections[key].pack(fill="both", expand=True)
        for k, btn in self._nav_btns.items():
            btn.activate() if k == key else btn.deactivate()

    # ── Утилиты ───────────────────────────────────────────────────────────────

    def _status(self, text: str, ok: bool = True):
        color, dot = (_SUCCESS, "●") if ok else (_ERROR, "⚠")
        self._status_dot.configure(text=f"{dot} {text}", text_color=color)
        self.update_idletasks()

    def _run(self, fn):
        threading.Thread(target=fn, daemon=True).start()

    # ── Виджеты-фабрики ───────────────────────────────────────────────────────

    @staticmethod
    def _title(parent, text: str, subtitle: str = ""):
        ctk.CTkLabel(
            parent, text=text, anchor="w",
            font=ctk.CTkFont(family=_FONT, size=26, weight="bold"),
            text_color=_TEXT1,
        ).pack(anchor="w", padx=32, pady=(32, 0))
        if subtitle:
            ctk.CTkLabel(
                parent, text=subtitle, anchor="w",
                font=ctk.CTkFont(family=_FONT, size=13),
                text_color=_TEXT2,
            ).pack(anchor="w", padx=32, pady=(4, 24))
        else:
            ctk.CTkFrame(parent, height=1, fg_color=_BORDER).pack(
                fill="x", padx=32, pady=(16, 24))

    @staticmethod
    def _card(parent, **kw) -> ctk.CTkFrame:
        return ctk.CTkFrame(
            parent, fg_color=_CARD, corner_radius=12,
            border_width=1, border_color=_BORDER, **kw)

    @staticmethod
    def _lbl(parent, text: str, secondary=False) -> ctk.CTkLabel:
        return ctk.CTkLabel(
            parent, text=text, anchor="w",
            font=ctk.CTkFont(family=_FONT, size=13),
            text_color=_TEXT2 if secondary else _TEXT1,
        )

    @staticmethod
    def _entry(parent, var=None, hint="", width=None) -> ctk.CTkEntry:
        kw = dict(
            corner_radius=8, border_width=1, border_color=_BORDER,
            fg_color="white", text_color=_TEXT1,
            placeholder_text=hint,
            font=ctk.CTkFont(family=_FONT, size=13),
        )
        if var is not None:
            kw["textvariable"] = var
        if width:
            kw["width"] = width
        return ctk.CTkEntry(parent, **kw)

    @staticmethod
    def _btn(parent, text: str, cmd, primary=True, width=None) -> ctk.CTkButton:
        kw = dict(
            text=text, command=cmd,
            corner_radius=8, height=36,
            font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
        )
        if primary:
            kw.update(fg_color=_ACCENT, hover_color=_ACCENT_H, text_color="white")
        else:
            kw.update(fg_color="#E8E8ED", hover_color="#D8D8DD", text_color=_TEXT1)
        if width:
            kw["width"] = width
        return ctk.CTkButton(parent, **kw)

    @staticmethod
    def _logbox(parent, height=170) -> ctk.CTkTextbox:
        return ctk.CTkTextbox(
            parent, height=height, corner_radius=8,
            border_width=1, border_color=_BORDER,
            fg_color="#F8F8FA", text_color=_TEXT1,
            font=ctk.CTkFont(family=_MONO, size=12),
            state="disabled", wrap="word",
        )

    @staticmethod
    def _append(box: ctk.CTkTextbox, msg: str):
        box.configure(state="normal")
        box.insert("end", msg + "\n")
        box.see("end")
        box.configure(state="disabled")

    @staticmethod
    def _combo(parent, var, width=260) -> ctk.CTkComboBox:
        return ctk.CTkComboBox(
            parent, variable=var, state="readonly", width=width,
            corner_radius=8, border_color=_BORDER, fg_color="white",
            button_color=_BORDER, button_hover_color=_TEXT2,
            font=ctk.CTkFont(family=_FONT, size=13),
            dropdown_font=ctk.CTkFont(family=_FONT, size=13),
        )

    # ── СЕКЦИЯ: Google Sheets ─────────────────────────────────────────────────

    def _build_sheets(self, f):
        self._title(f, "Google Sheets", "Загрузка данных из таблицы")

        card = self._card(f)
        card.pack(fill="x", padx=32, pady=(0, 16))

        self._lbl(card, "Ссылка на таблицу").pack(anchor="w", padx=20, pady=(18, 4))
        self._sheet_url = tk.StringVar()
        self._entry(card, var=self._sheet_url,
                    hint="https://docs.google.com/spreadsheets/d/…"
                    ).pack(fill="x", padx=20, pady=(0, 14))

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=(0, 18))
        self._lbl(row, "Диапазон:").pack(side="left")
        self._sheet_range = tk.StringVar()
        self._entry(row, var=self._sheet_range, hint="Лист1!A:Z", width=200
                    ).pack(side="left", padx=(10, 10))
        self._lbl(row, "(необязательно)", secondary=True).pack(side="left")

        self._btn_load = self._btn(card, "  Загрузить данные  ", self._load_sheet, width=190)
        self._btn_load.pack(anchor="w", padx=20, pady=(0, 20))

        self._lbl(f, "Журнал").pack(anchor="w", padx=32, pady=(0, 6))
        self._log_sheets = self._logbox(f, height=160)
        self._log_sheets.pack(fill="x", padx=32, pady=(0, 18))

        self._lbl(f, "Загруженные таблицы").pack(anchor="w", padx=32, pady=(0, 6))
        self._sheets_list = ctk.CTkScrollableFrame(
            f, height=110, corner_radius=10,
            fg_color=_CARD, border_width=1, border_color=_BORDER,
        )
        self._sheets_list.pack(fill="x", padx=32, pady=(0, 32))
        self._refresh_sheets_list()

    def _load_sheet(self):
        url = self._sheet_url.get().strip()
        if not url:
            messagebox.showwarning("Внимание", "Вставьте ссылку на таблицу.")
            return
        if not extract_sheet_id(url):
            messagebox.showerror("Ошибка", "Некорректная ссылка на Google Sheets.")
            return
        self._btn_load.configure(state="disabled")
        self._status("Загрузка данных…")
        self._append(self._log_sheets, f"→ {url}")

        def task():
            try:
                rng = self._sheet_range.get().strip()
                tbl, cnt = sync_sheet_to_database(url, rng, self.db)
                self.after(0, lambda: self._sheet_ok(tbl, cnt))
            except Exception as e:
                self.after(0, lambda m=str(e): self._sheet_err(m))

        self._run(task)

    def _sheet_ok(self, tbl, cnt):
        self._btn_load.configure(state="normal")
        self._status("Готово")
        self._append(self._log_sheets, f"✓ Загружено: «{tbl}», строк: {cnt}")
        self._refresh_sheets_list()
        messagebox.showinfo("Успех", f"Таблица: {tbl}\nСтрок: {cnt}")

    def _sheet_err(self, err):
        self._btn_load.configure(state="normal")
        self._status("Ошибка", ok=False)
        self._append(self._log_sheets, f"✗ {err}")
        messagebox.showerror("Ошибка загрузки", err)

    def _refresh_sheets_list(self):
        for w in self._sheets_list.winfo_children():
            w.destroy()
        sheets = self.db.get_all_sheets()
        if not sheets:
            self._lbl(self._sheets_list, "Нет загруженных таблиц", secondary=True
                      ).pack(padx=16, pady=12)
            return
        for s in sheets:
            row = ctk.CTkFrame(self._sheets_list, fg_color="transparent")
            row.pack(fill="x", pady=4, padx=12)
            ctk.CTkLabel(
                row, text=s["sheet_name"] or "Без названия",
                font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
                text_color=_TEXT1, anchor="w",
            ).pack(side="left")
            ctk.CTkLabel(
                row, text=f"  ·  {s['row_count']} строк",
                font=ctk.CTkFont(family=_FONT, size=12),
                text_color=_TEXT2, anchor="w",
            ).pack(side="left")

    # ── СЕКЦИЯ: Шаблоны ───────────────────────────────────────────────────────

    def _build_templates(self, f):
        self._title(f, "Шаблоны", "Скачивание и генерация документов")

        # Карточка: скачать шаблон
        c1 = self._card(f)
        c1.pack(fill="x", padx=32, pady=(0, 16))

        self._lbl(c1, "Ссылка на шаблон Google Docs").pack(
            anchor="w", padx=20, pady=(18, 4))
        self._doc_url = tk.StringVar()
        self._entry(c1, var=self._doc_url,
                    hint="https://docs.google.com/document/d/…"
                    ).pack(fill="x", padx=20, pady=(0, 14))
        self._btn_dl = self._btn(c1, "  Скачать шаблон  ",
                                 self._download_template, width=180)
        self._btn_dl.pack(anchor="w", padx=20, pady=(0, 20))

        # Карточка: генерация
        c2 = self._card(f)
        c2.pack(fill="x", padx=32, pady=(0, 16))

        self._lbl(c2, "Таблица с данными:").pack(anchor="w", padx=20, pady=(18, 4))
        self._tpl_table = tk.StringVar()
        self._tpl_table.trace_add("write", lambda *_: self._refresh_group_by_cols())
        self._tbl_picker = _TablePickerWidget(c2, self.db, self._tpl_table)
        self._tbl_picker.pack(fill="x", padx=20, pady=(0, 6))
        r1_btn = ctk.CTkFrame(c2, fg_color="transparent")
        r1_btn.pack(fill="x", padx=20, pady=(0, 10))
        self._btn(r1_btn, "Обновить таблицы", self._refresh_tpl_tables,
                  primary=False, width=140).pack(side="left")

        r2 = ctk.CTkFrame(c2, fg_color="transparent")
        r2.pack(fill="x", padx=20, pady=(0, 10))
        self._lbl(r2, "Столбец для имён файлов:").pack(side="left")
        self._filename_col = tk.StringVar()
        self._entry(r2, var=self._filename_col,
                    hint="необязательно", width=200
                    ).pack(side="left", padx=(10, 0))

        r_num = ctk.CTkFrame(c2, fg_color="transparent")
        r_num.pack(fill="x", padx=20, pady=(0, 14))
        self._lbl(r_num, "Нумерация строк  {{ s.").pack(side="left")
        self._num_var = tk.StringVar(value="num")
        self._entry(r_num, var=self._num_var, width=90
                    ).pack(side="left", padx=(4, 0))
        self._lbl(r_num, " }}").pack(side="left")
        self._lbl(r_num, "   ← оставьте пустым чтобы не нумеровать",
                  secondary=True).pack(side="left", padx=(8, 0))

        # ── Режим единого документа ──
        r_single = ctk.CTkFrame(c2, fg_color="transparent")
        r_single.pack(fill="x", padx=20, pady=(0, 10))
        self._single_doc = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            r_single, text="Создать один документ для всех строк",
            variable=self._single_doc,
            corner_radius=4, fg_color=_ACCENT, hover_color=_ACCENT_H,
            font=ctk.CTkFont(family=_FONT, size=13), text_color=_TEXT1,
        ).pack(side="left", padx=(0, 16))
        self._lbl(r_single, "Имя файла:").pack(side="left")
        self._single_doc_name = tk.StringVar(value="document.docx")
        self._entry(r_single, var=self._single_doc_name, width=180
                    ).pack(side="left", padx=(6, 0))

        # ── Многоуровневая группировка ──
        self._lbl(c2, "Уровни группировки:").pack(
            anchor="w", padx=20, pady=(0, 4))

        # Список текущих уровней: [(col_name, label_tpl_var), ...]
        self._group_levels: list[tuple[str, tk.StringVar]] = []

        self._group_levels_frame = ctk.CTkFrame(
            c2, fg_color="#F5F5F7", corner_radius=8,
            border_width=1, border_color=_BORDER)
        self._group_levels_frame.pack(fill="x", padx=20, pady=(0, 6))
        self._rebuild_group_levels_ui()

        r_add = ctk.CTkFrame(c2, fg_color="transparent")
        r_add.pack(fill="x", padx=20, pady=(0, 14))
        self._add_group_col = tk.StringVar()
        self._add_group_combo = self._combo(r_add, self._add_group_col, width=200)
        self._add_group_combo.pack(side="left", padx=(0, 6))
        self._btn(r_add, "+ Добавить", self._add_group_level,
                  primary=False, width=110).pack(side="left", padx=(0, 6))
        self._btn(r_add, "Обновить список", self._refresh_group_by_cols,
                  primary=False, width=130).pack(side="left")

        self._merge_left = tk.BooleanVar(value=True)
        self._merge_right = tk.BooleanVar(value=False)
        self._merge_left.trace_add("write", self._on_merge_left_changed)
        self._merge_right.trace_add("write", self._on_merge_right_changed)

        r_merge = ctk.CTkFrame(c2, fg_color="transparent")
        r_merge.pack(anchor="w", padx=20, pady=(0, 14))
        ctk.CTkCheckBox(
            r_merge, text="Объединять пустые ячейки влево",
            variable=self._merge_left,
            corner_radius=4, fg_color=_ACCENT, hover_color=_ACCENT_H,
            font=ctk.CTkFont(family=_FONT, size=13), text_color=_TEXT1,
        ).pack(side="left", padx=(0, 24))
        ctk.CTkCheckBox(
            r_merge, text="Объединять пустые ячейки вправо",
            variable=self._merge_right,
            corner_radius=4, fg_color=_ACCENT, hover_color=_ACCENT_H,
            font=ctk.CTkFont(family=_FONT, size=13), text_color=_TEXT1,
        ).pack(side="left")

        # ── Правила слияния по столбцам ──
        self._lbl(c2, "Правила слияния столбцов:").pack(
            anchor="w", padx=20, pady=(0, 4))

        self._col_merge_rules: list[tuple[str, tk.StringVar]] = []

        self._col_merge_frame = ctk.CTkFrame(
            c2, fg_color="#F5F5F7", corner_radius=8,
            border_width=1, border_color=_BORDER)
        self._col_merge_frame.pack(fill="x", padx=20, pady=(0, 6))
        self._rebuild_col_merge_ui()

        r_col_add = ctk.CTkFrame(c2, fg_color="transparent")
        r_col_add.pack(fill="x", padx=20, pady=(0, 14))
        self._add_cmr_col = tk.StringVar()
        self._add_cmr_combo = self._combo(r_col_add, self._add_cmr_col, width=200)
        self._add_cmr_combo.pack(side="left", padx=(0, 6))
        self._add_cmr_dir = tk.StringVar(value="влево")
        ctk.CTkComboBox(
            r_col_add, variable=self._add_cmr_dir,
            values=["влево", "вправо"], state="readonly", width=100,
            corner_radius=8, border_color=_BORDER, fg_color="white",
            button_color=_BORDER, button_hover_color=_TEXT2,
            font=ctk.CTkFont(family=_FONT, size=13),
            dropdown_font=ctk.CTkFont(family=_FONT, size=13),
        ).pack(side="left", padx=(0, 6))
        self._btn(r_col_add, "+ Добавить", self._add_col_merge_rule,
                  primary=False, width=110).pack(side="left")

        self._btn_gen = self._btn(c2, "  Сгенерировать документы  ",
                                  self._generate_docs, width=230)
        self._btn_gen.pack(anchor="w", padx=20, pady=(4, 20))

        self._lbl(f, "Журнал").pack(anchor="w", padx=32, pady=(0, 6))
        self._log_tpl = self._logbox(f, height=220)
        self._log_tpl.pack(fill="x", padx=32, pady=(0, 32))

        self._refresh_tpl_tables()

    def _refresh_tpl_tables(self):
        tables_info = self.db.get_tables_with_display_names()
        internals = [t[0] for t in tables_info]
        if hasattr(self, "_tbl_picker"):
            self._tbl_picker.refresh(tables_info)
        if internals and self._tpl_table.get() not in internals:
            self._tpl_table.set(internals[0])
        if hasattr(self, "_db_combo"):
            self._refresh_db()
        self._refresh_group_by_cols()

    def _on_merge_left_changed(self, *_):
        if self._merge_left.get():
            self._merge_right.set(False)

    def _on_merge_right_changed(self, *_):
        if self._merge_right.get():
            self._merge_left.set(False)

    def _refresh_group_by_cols(self):
        """Обновляет списки колонок для группировки и нумерации."""
        tbl = self._tpl_table.get()
        cols = [""]
        if tbl:
            try:
                orig = self.db.get_original_headers(tbl)
                cols.extend(sorted(orig.keys()))
            except Exception:
                pass
        if hasattr(self, "_add_group_combo"):
            self._add_group_combo.configure(values=cols)
        if hasattr(self, "_add_cmr_combo"):
            self._add_cmr_combo.configure(values=cols)

    def _add_group_level(self):
        """Добавляет выбранную колонку в список уровней группировки."""
        col = self._add_group_col.get().strip()
        if not col:
            return
        if any(c == col for c, _ in self._group_levels):
            return  # дубликаты не добавляем
        label_var = tk.StringVar(value="{value}")
        self._group_levels.append((col, label_var))
        self._rebuild_group_levels_ui()

    def _remove_group_level(self, idx: int):
        """Удаляет уровень группировки по индексу."""
        if 0 <= idx < len(self._group_levels):
            self._group_levels.pop(idx)
            self._rebuild_group_levels_ui()

    def _rebuild_group_levels_ui(self):
        """Перестраивает виджет списка уровней группировки."""
        for w in self._group_levels_frame.winfo_children():
            w.destroy()

        if not self._group_levels:
            self._lbl(self._group_levels_frame,
                      "Группировка не задана — все строки будут добавлены без разбивки",
                      secondary=True).pack(padx=12, pady=8)
            return

        for i, (col, label_var) in enumerate(self._group_levels):
            row = ctk.CTkFrame(self._group_levels_frame, fg_color="transparent")
            row.pack(fill="x", padx=8, pady=4)

            ctk.CTkLabel(
                row, text=f"{i + 1}.", width=22, anchor="e",
                font=ctk.CTkFont(family=_FONT, size=12), text_color=_TEXT2,
            ).pack(side="left")
            ctk.CTkLabel(
                row, text=col, width=150, anchor="w",
                font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
                text_color=_TEXT1,
            ).pack(side="left", padx=(4, 8))
            self._entry(row, var=label_var,
                        hint="напр. «Группа {value}»", width=180
                        ).pack(side="left", padx=(0, 6))
            self._btn(row, "×",
                      lambda idx=i: self._remove_group_level(idx),
                      primary=False, width=32).pack(side="left")

    def _rebuild_col_merge_ui(self):
        for w in self._col_merge_frame.winfo_children():
            w.destroy()
        if not self._col_merge_rules:
            self._lbl(self._col_merge_frame,
                      "Нет правил — используются глобальные настройки слияния",
                      secondary=True).pack(padx=12, pady=8)
            return
        for i, (col, dir_var) in enumerate(self._col_merge_rules):
            row = ctk.CTkFrame(self._col_merge_frame, fg_color="transparent")
            row.pack(fill="x", padx=8, pady=4)
            ctk.CTkLabel(
                row, text=f"{i + 1}.", width=22, anchor="e",
                font=ctk.CTkFont(family=_FONT, size=12), text_color=_TEXT2,
            ).pack(side="left")
            ctk.CTkLabel(
                row, text=col, width=150, anchor="w",
                font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
                text_color=_TEXT1,
            ).pack(side="left", padx=(4, 8))
            ctk.CTkComboBox(
                row, variable=dir_var, values=["влево", "вправо"],
                state="readonly", width=100,
                corner_radius=8, border_color=_BORDER, fg_color="white",
                button_color=_BORDER, button_hover_color=_TEXT2,
                font=ctk.CTkFont(family=_FONT, size=13),
                dropdown_font=ctk.CTkFont(family=_FONT, size=13),
            ).pack(side="left", padx=(0, 6))
            self._btn(row, "×",
                      lambda idx=i: self._remove_col_merge_rule(idx),
                      primary=False, width=32).pack(side="left")

    def _add_col_merge_rule(self):
        col = self._add_cmr_col.get().strip()
        if not col:
            return
        if any(c == col for c, _ in self._col_merge_rules):
            return
        dir_var = tk.StringVar(value=self._add_cmr_dir.get())
        self._col_merge_rules.append((col, dir_var))
        self._rebuild_col_merge_ui()

    def _remove_col_merge_rule(self, idx: int):
        if 0 <= idx < len(self._col_merge_rules):
            self._col_merge_rules.pop(idx)
            self._rebuild_col_merge_ui()

    def _download_template(self):
        url = self._doc_url.get().strip()
        if not url:
            messagebox.showwarning("Внимание", "Вставьте ссылку на шаблон.")
            return
        if not extract_doc_id(url):
            messagebox.showerror("Ошибка", "Некорректная ссылка на Google Docs.")
            return
        self._btn_dl.configure(state="disabled")
        self._status("Скачивание шаблона…")
        self._append(self._log_tpl, f"→ Скачиваю: {url}")

        def task():
            try:
                path = download_template(url, self.db)
                ph = find_placeholders(path)
                self.after(0, lambda: self._dl_ok(path, ph))
            except Exception as e:
                self.after(0, lambda m=str(e): self._dl_err(m))

        self._run(task)

    def _dl_ok(self, path, placeholders):
        self._btn_dl.configure(state="normal")
        self._status("Готово")
        self._append(self._log_tpl, f"✓ Сохранён: {path.name}")
        self._append(self._log_tpl,
                     f"  Плейсхолдеры: {', '.join(placeholders) or 'не найдены'}")

    def _dl_err(self, err):
        self._btn_dl.configure(state="normal")
        self._status("Ошибка", ok=False)
        self._append(self._log_tpl, f"✗ {err}")
        messagebox.showerror("Ошибка скачивания", err)

    def _generate_docs(self):
        tbl = self._tpl_table.get()
        if not tbl:
            messagebox.showwarning("Внимание", "Выберите таблицу с данными.")
            return
        templates = self.db.get_all_templates()
        if not templates:
            messagebox.showwarning("Внимание", "Сначала скачайте шаблон.")
            return
        tpl_path = Path(templates[0]["local_path"])
        if not tpl_path.exists():
            messagebox.showerror("Ошибка", f"Файл шаблона не найден:\n{tpl_path}")
            return
        self._btn_gen.configure(state="disabled")
        self._status("Генерация документов…")

        single = self._single_doc.get()

        def task():
            try:
                group_by = [c for c, _ in self._group_levels] or None
                label_tpls = [
                    v.get().strip() or "{value}"
                    for _, v in self._group_levels
                ] or None
                merge_left = self._merge_left.get()
                merge_right = self._merge_right.get()
                num_var = self._num_var.get().strip() or None

                column_merge_rules = None
                if self._col_merge_rules:
                    column_merge_rules = [
                        (col, "left" if dir_var.get() == "влево" else "right")
                        for col, dir_var in self._col_merge_rules
                    ]

                if single:
                    headers, all_rows = self.db.get_all_data(tbl)
                    orig_map = self.db.get_original_headers(tbl)
                    ts = datetime.now().strftime("%H.%M %d.%m.%Y")
                    raw_name = self._single_doc_name.get().strip() or "document"
                    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", raw_name)
                    base = safe_name[:-5] if safe_name.lower().endswith(".docx") else safe_name
                    safe_name = f"{base} ({ts}).docx"
                    path = generate_document(
                        template_path=tpl_path,
                        row_data={},
                        output_filename=safe_name,
                        all_rows_data=(headers, all_rows),
                        original_headers_map=orig_map,
                        group_by=group_by,
                        group_label_templates=label_tpls,
                        merge_empty_cells=merge_left or merge_right,
                        merge_right=merge_right,
                        num_var=num_var,
                        column_merge_rules=column_merge_rules,
                    )
                    files = [path]
                else:
                    col = self._filename_col.get().strip() or None
                    files = generate_documents_for_all_rows(
                        template_path=tpl_path,
                        table_name=tbl,
                        filename_column=col,
                        db=self.db,
                        group_by=group_by,
                        group_label_templates=label_tpls,
                        merge_empty_cells=merge_left or merge_right,
                        merge_right=merge_right,
                        num_var=num_var,
                        column_merge_rules=column_merge_rules,
                    )
                self.after(0, lambda: self._gen_ok(files))
            except Exception as e:
                self.after(0, lambda m=str(e): self._gen_err(m))

        self._run(task)

    def _gen_ok(self, files):
        self._btn_gen.configure(state="normal")
        self._status("Готово")
        self._append(self._log_tpl, f"✓ Сгенерировано: {len(files)} документов")
        for f in files:
            self._append(self._log_tpl, f"    {f.name}")
        messagebox.showinfo("Успех",
                            f"Сгенерировано {len(files)} документов.\n"
                            f"Папка: {config.OUTPUT_DIR}")

    def _gen_err(self, err):
        self._btn_gen.configure(state="normal")
        self._status("Ошибка генерации", ok=False)
        self._append(self._log_tpl, f"✗ {err}")
        messagebox.showerror("Ошибка генерации", err)

    # ── СЕКЦИЯ: Почта ─────────────────────────────────────────────────────────

    def _build_email(self, f):
        self._title(f, "Почта", "Отправка писем студентам")

        card = self._card(f)
        card.pack(fill="x", padx=32, pady=(0, 16))

        self._lbl(card, "Получатель (несколько — через ;)").pack(
            anchor="w", padx=20, pady=(18, 4))
        self._email_to = tk.StringVar()
        self._entry(card, var=self._email_to,
                    hint="user@example.com; user2@example.com"
                    ).pack(fill="x", padx=20, pady=(0, 14))

        self._lbl(card, "Тема письма").pack(anchor="w", padx=20, pady=(0, 4))
        self._email_subj = tk.StringVar()
        self._entry(card, var=self._email_subj, hint="Тема"
                    ).pack(fill="x", padx=20, pady=(0, 14))

        self._lbl(card, "Текст письма").pack(anchor="w", padx=20, pady=(0, 4))
        self._email_body = ctk.CTkTextbox(
            card, height=110, corner_radius=8,
            border_width=1, border_color=_BORDER,
            fg_color="white", text_color=_TEXT1,
            font=ctk.CTkFont(family=_FONT, size=13),
        )
        self._email_body.pack(fill="x", padx=20, pady=(0, 14))

        # Вложение
        ar = ctk.CTkFrame(card, fg_color="transparent")
        ar.pack(fill="x", padx=20, pady=(0, 14))
        self._lbl(ar, "Вложение:").pack(side="left")
        self._attach = tk.StringVar()
        self._entry(ar, var=self._attach, hint="путь к файлу", width=330
                    ).pack(side="left", padx=(10, 6))
        self._btn(ar, "Выбрать…", self._pick_attach, primary=False, width=100
                  ).pack(side="left")

        # Тестовый режим
        self._dry_run = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            card, text="Тестовый режим (письма не отправляются)",
            variable=self._dry_run,
            corner_radius=4, fg_color=_ACCENT, hover_color=_ACCENT_H,
            font=ctk.CTkFont(family=_FONT, size=13), text_color=_TEXT1,
        ).pack(anchor="w", padx=20, pady=(0, 14))

        self._btn_send = self._btn(card, "  Отправить  ", self._send_email, width=150)
        self._btn_send.pack(anchor="w", padx=20, pady=(0, 20))

        self._lbl(f, "Журнал отправки").pack(anchor="w", padx=32, pady=(0, 6))
        self._log_email = self._logbox(f, height=200)
        self._log_email.pack(fill="x", padx=32, pady=(0, 32))

    def _pick_attach(self):
        p = filedialog.askopenfilename(
            title="Выберите файл",
            filetypes=[("Word", "*.docx"), ("PDF", "*.pdf"), ("Все", "*.*")],
        )
        if p:
            self._attach.set(p)

    def _send_email(self):
        to = self._email_to.get().strip()
        subj = self._email_subj.get().strip()
        body = self._email_body.get("1.0", "end").strip()
        if not to:
            messagebox.showwarning("Внимание", "Укажите получателя.")
            return
        if not subj:
            messagebox.showwarning("Внимание", "Укажите тему.")
            return
        if not body:
            messagebox.showwarning("Внимание", "Введите текст письма.")
            return
        self._btn_send.configure(state="disabled")
        self._status("Отправка письма…")

        def task():
            try:
                sender = EmailSender(db=self.db, dry_run=self._dry_run.get())
                ap = self._attach.get().strip()
                attach_path = Path(ap) if ap else None
                recipients = [r.strip() for r in to.split(";") if r.strip()]
                res = {"sent": 0, "failed": 0, "errors": []}
                for r in recipients:
                    try:
                        sender.send_email(
                            to=r, subject=subj, body=body,
                            attachment_path=attach_path)
                        res["sent"] += 1
                    except Exception as e:
                        res["failed"] += 1
                        res["errors"].append(f"{r}: {e}")
                self.after(0, lambda: self._mail_ok(res))
            except Exception as e:
                self.after(0, lambda m=str(e): self._mail_err(m))

        self._run(task)

    def _mail_ok(self, res):
        self._btn_send.configure(state="normal")
        self._status("Готово")
        mode = "ТЕСТ" if self._dry_run.get() else "Отправлено"
        self._append(self._log_email,
                     f"[{mode}] ✓ {res['sent']}  ✗ {res['failed']}")
        for err in res["errors"]:
            self._append(self._log_email, f"  ✗ {err}")
        if res["failed"] == 0:
            messagebox.showinfo("Результат", f"Успешно: {res['sent']}")
        else:
            messagebox.showwarning("Результат",
                                   f"Успешно: {res['sent']}\nОшибок: {res['failed']}")

    def _mail_err(self, err):
        self._btn_send.configure(state="normal")
        self._status("Ошибка отправки", ok=False)
        self._append(self._log_email, f"✗ {err}")
        messagebox.showerror("Ошибка", err)

    # ── СЕКЦИЯ: База данных ───────────────────────────────────────────────────

    def _build_database(self, f):
        self._title(f, "База данных", "Просмотр загруженных данных")

        ctrl = self._card(f)
        ctrl.pack(fill="x", padx=32, pady=(0, 16))

        r = ctk.CTkFrame(ctrl, fg_color="transparent")
        r.pack(fill="x", padx=20, pady=(18, 6))
        self._lbl(r, "Таблица:").pack(side="left")
        self._db_tbl = tk.StringVar()
        self._db_combo = self._combo(r, self._db_tbl, width=280)
        self._db_combo.pack(side="left", padx=(10, 6))
        self._btn(r, "Обновить список", self._refresh_db, primary=False, width=130
                  ).pack(side="left", padx=(0, 6))
        self._btn(r, "Показать данные", self._show_data, width=140
                  ).pack(side="left")

        r2 = ctk.CTkFrame(ctrl, fg_color="transparent")
        r2.pack(fill="x", padx=20, pady=(0, 18))
        self._btn(r2, "Переименовать таблицу", self._open_table_rename_dialog,
                  primary=False, width=180).pack(side="left", padx=(0, 6))
        self._btn(r2, "Переименовать столбцы", self._open_rename_dialog,
                  primary=False, width=180).pack(side="left", padx=(0, 6))
        self._btn(r2, "Объединить столбцы", self._open_merge_dialog,
                  primary=False, width=160).pack(side="left")

        # Treeview
        tree_card = self._card(f)
        tree_card.pack(fill="both", expand=True, padx=32, pady=(0, 8))

        tree_wrap = tk.Frame(tree_card, bg=_CARD)
        tree_wrap.pack(fill="both", expand=True, padx=12, pady=12)

        style = tkttk.Style()
        style.theme_use("default")
        style.configure(
            "Mac.Treeview",
            background="white", fieldbackground="white",
            foreground=_TEXT1, rowheight=30,
            font=(_FONT, 12), borderwidth=0,
        )
        style.configure(
            "Mac.Treeview.Heading",
            background=_BG, foreground=_TEXT2,
            font=(_FONT, 12, "bold"), relief="flat",
        )
        style.map(
            "Mac.Treeview",
            background=[("selected", "#E3EEFB")],
            foreground=[("selected", _TEXT1)],
        )

        self._tree = tkttk.Treeview(
            tree_wrap, show="headings", style="Mac.Treeview")
        vsb = tkttk.Scrollbar(tree_wrap, orient="vertical",
                               command=self._tree.yview)
        hsb = tkttk.Scrollbar(tree_wrap, orient="horizontal",
                               command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_wrap.grid_rowconfigure(0, weight=1)
        tree_wrap.grid_columnconfigure(0, weight=1)

        self._db_info = tk.StringVar()
        ctk.CTkLabel(
            f, textvariable=self._db_info, anchor="w",
            font=ctk.CTkFont(family=_FONT, size=12), text_color=_TEXT2,
        ).pack(anchor="w", padx=32, pady=(4, 32))

        self._db_name_map: dict[str, str] = {}
        self._refresh_db()

    def _refresh_db(self):
        tables_info = self.db.get_tables_with_display_names()
        display_names = [d for _, d in tables_info]
        self._db_name_map = {d: t for t, d in tables_info}
        self._db_combo.configure(values=display_names)
        cur = self._db_tbl.get()
        if display_names and cur not in display_names:
            self._db_tbl.set(display_names[0])

    def _show_data(self):
        tbl = getattr(self, "_db_name_map", {}).get(self._db_tbl.get(), self._db_tbl.get())
        if not tbl:
            messagebox.showwarning("Внимание", "Выберите таблицу.")
            return
        try:
            headers, rows = self.db.get_all_data(tbl)
            orig = self.db.get_original_headers(tbl)
            aliases = self.db.get_column_aliases(tbl)
            rev = {v: k for k, v in orig.items()}

            self._tree.delete(*self._tree.get_children())
            self._tree["columns"] = headers
            for col in headers:
                display = aliases.get(col) or rev.get(col, col)
                self._tree.heading(
                    col, text=display,
                    command=lambda c=col, t=tbl: self._rename_column_quick(c, t),
                )
                self._tree.column(col, width=140, minwidth=80)
            for i, row in enumerate(rows):
                tag = "odd" if i % 2 else "even"
                self._tree.insert("", "end", values=row, tags=(tag,))
            self._tree.tag_configure("odd", background=_ODD_ROW)
            self._tree.tag_configure("even", background="white")

            self._db_info.set(
                f"Таблица: {tbl}   ·   "
                f"Столбцов: {len(headers)}   ·   "
                f"Строк: {len(rows)}   ·   "
                f"Нажмите на заголовок столбца чтобы переименовать"
            )
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные:\n{e}")

    def _open_rename_dialog(self):
        tbl = getattr(self, "_db_name_map", {}).get(self._db_tbl.get(), self._db_tbl.get())
        if not tbl:
            messagebox.showwarning("Внимание", "Выберите таблицу.")
            return
        _ColumnRenameDialog(self, self.db, tbl, on_save=self._show_data)

    def _open_merge_dialog(self):
        tbl = getattr(self, "_db_name_map", {}).get(self._db_tbl.get(), self._db_tbl.get())
        if not tbl:
            messagebox.showwarning("Внимание", "Выберите таблицу.")
            return
        _MergeColumnsDialog(self, self.db, tbl, on_merge=self._show_data)

    def _rename_column_quick(self, safe_col: str, table_name: str):
        """Открывает мини-диалог переименования одного столбца по клику на заголовок."""
        orig = self.db.get_original_headers(table_name)
        aliases = self.db.get_column_aliases(table_name)
        rev = {v: k for k, v in orig.items()}
        orig_name = rev.get(safe_col, safe_col)
        current_alias = aliases.get(safe_col, "")

        dialog = ctk.CTkToplevel(self)
        dialog.title("Переименовать столбец")
        dialog.geometry("420x210")
        dialog.resizable(False, False)
        dialog.configure(fg_color=_BG)
        dialog.after(150, lambda: dialog.grab_set() if dialog.winfo_exists() else None)

        ctk.CTkLabel(
            dialog, text=f"Исходное имя: {orig_name}",
            font=ctk.CTkFont(family=_FONT, size=12), text_color=_TEXT2,
        ).pack(anchor="w", padx=24, pady=(20, 4))
        ctk.CTkLabel(
            dialog, text="Локальное имя:",
            font=ctk.CTkFont(family=_FONT, size=13), text_color=_TEXT1,
        ).pack(anchor="w", padx=24, pady=(0, 4))

        var = tk.StringVar(value=current_alias)
        entry = ctk.CTkEntry(
            dialog, textvariable=var, width=370,
            corner_radius=8, border_width=1, border_color=_BORDER,
            fg_color="white", text_color=_TEXT1,
            placeholder_text=orig_name,
            font=ctk.CTkFont(family=_FONT, size=13),
        )
        entry.pack(padx=24, pady=(0, 16))
        entry.focus_set()

        def _save():
            self.db.set_column_alias(table_name, safe_col, var.get())
            dialog.destroy()
            self._show_data()

        btn_row = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_row.pack(fill="x", padx=24)
        ctk.CTkButton(
            btn_row, text="Отмена", command=dialog.destroy,
            corner_radius=8, height=36, width=100,
            fg_color="#E8E8ED", hover_color="#D8D8DD", text_color=_TEXT1,
            font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
        ).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            btn_row, text="Сохранить", command=_save,
            corner_radius=8, height=36, width=110,
            fg_color=_ACCENT, hover_color=_ACCENT_H, text_color="white",
            font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
        ).pack(side="right")
        dialog.bind("<Return>", lambda e: _save())

    def _open_table_rename_dialog(self):
        """Диалог переименования выбранной таблицы."""
        disp = self._db_tbl.get()
        if not disp:
            messagebox.showwarning("Внимание", "Выберите таблицу.")
            return
        tbl = getattr(self, "_db_name_map", {}).get(disp, disp)

        dialog = ctk.CTkToplevel(self)
        dialog.title("Переименовать таблицу")
        dialog.geometry("440x210")
        dialog.resizable(False, False)
        dialog.configure(fg_color=_BG)
        dialog.after(150, lambda: dialog.grab_set() if dialog.winfo_exists() else None)

        ctk.CTkLabel(
            dialog, text=f"Внутреннее имя: {tbl}",
            font=ctk.CTkFont(family=_FONT, size=12), text_color=_TEXT2,
        ).pack(anchor="w", padx=24, pady=(20, 4))
        ctk.CTkLabel(
            dialog, text="Отображаемое имя:",
            font=ctk.CTkFont(family=_FONT, size=13), text_color=_TEXT1,
        ).pack(anchor="w", padx=24, pady=(0, 4))

        # Pre-fill if user already set a custom name
        prefill = disp if disp != tbl else ""
        var = tk.StringVar(value=prefill)
        entry = ctk.CTkEntry(
            dialog, textvariable=var, width=390,
            corner_radius=8, border_width=1, border_color=_BORDER,
            fg_color="white", text_color=_TEXT1,
            placeholder_text="Введите имя таблицы",
            font=ctk.CTkFont(family=_FONT, size=13),
        )
        entry.pack(padx=24, pady=(0, 16))
        entry.focus_set()

        def _save():
            self.db.set_table_display_name(tbl, var.get())
            dialog.destroy()
            self._refresh_db()
            if hasattr(self, "_tbl_picker"):
                self._refresh_tpl_tables()

        btn_row = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_row.pack(fill="x", padx=24)
        ctk.CTkButton(
            btn_row, text="Отмена", command=dialog.destroy,
            corner_radius=8, height=36, width=100,
            fg_color="#E8E8ED", hover_color="#D8D8DD", text_color=_TEXT1,
            font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
        ).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            btn_row, text="Сохранить", command=_save,
            corner_radius=8, height=36, width=110,
            fg_color=_ACCENT, hover_color=_ACCENT_H, text_color="white",
            font=ctk.CTkFont(family=_FONT, size=13, weight="bold"),
        ).pack(side="right")
        dialog.bind("<Return>", lambda e: _save())

    # ── Закрытие ──────────────────────────────────────────────────────────────

    def _on_close(self):
        self.db.close()
        self.destroy()

# -*- coding: utf-8 -*-
# log_ui.py — компактный UI-модуль лога для Tkinter

from __future__ import annotations
from typing import Tuple, Optional
import tkinter as tk
from tkinter import ttk, filedialog

_TXT: Optional[tk.Text] = None

def _bind_text_shortcuts(w: tk.Text):
    w.bind("<Control-a>", lambda e: (w.tag_add("sel", "1.0", "end-1c"), "break"))
    w.bind("<Control-A>", lambda e: (w.tag_add("sel", "1.0", "end-1c"), "break"))
    w.bind("<Control-c>", lambda e: (w.event_generate("<<Copy>>"), "break"))
    w.bind("<Control-C>", lambda e: (w.event_generate("<<Copy>>"), "break"))
    w.bind("<Control-v>", lambda e: (w.event_generate("<<Paste>>"), "break"))
    w.bind("<Control-V>", lambda e: (w.event_generate("<<Paste>>"), "break"))

def _save_to_file():
    if _TXT is None:
        return
    try:
        p = filedialog.asksaveasfilename(
            title="Сохранить лог",
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("All files", "*.*")]
        )
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write(_TXT.get("1.0", "end-1c"))
    except Exception:
        pass

def create_log_panel(parent, *, height: int = 6) -> Tuple[tk.Frame, tk.Text]:
    """
    Создаёт панель лога (Frame) с текстовым полем + скроллбаром и контекстным меню.
    Возвращает (frame, text_widget). Вызов .pack/.grid делается снаружи.
    """
    global _TXT

    frm = ttk.Frame(parent)

    txt = tk.Text(frm, height=height, wrap="word", undo=True)
    vsb = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
    txt.configure(yscrollcommand=vsb.set)

    txt.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    frm.rowconfigure(0, weight=1)
    frm.columnconfigure(0, weight=1)

    _bind_text_shortcuts(txt)

    # контекстное меню
    menu = tk.Menu(txt, tearoff=0)
    menu.add_command(label="Копировать", command=lambda: txt.event_generate("<<Copy>>"))
    menu.add_command(label="Вставить",   command=lambda: txt.event_generate("<<Paste>>"))
    menu.add_separator()
    menu.add_command(label="Выделить всё", command=lambda: txt.tag_add("sel", "1.0", "end-1c"))
    menu.add_command(label="Очистить",     command=lambda: txt.delete("1.0", "end"))
    menu.add_separator()
    menu.add_command(label="Сохранить в файл…", command=_save_to_file)

    def _show_menu(e):
        try:
            menu.tk_popup(e.x_root, e.y_root)
        finally:
            menu.grab_release()

    txt.bind("<Button-3>", _show_menu)

    _TXT = txt
    return frm, txt

def log(msg: str):
    """Пишет строку в лог и прокручивает к концу (безопасно для пустого состояния)."""
    if _TXT is None:
        return
    try:
        _TXT.insert("end", str(msg).rstrip() + "\n")
        _TXT.see("end")
    except Exception:
        pass

def clear():
    if _TXT is None:
        return
    _TXT.delete("1.0", "end")

def get_text() -> str:
    if _TXT is None:
        return ""
    return _TXT.get("1.0", "end-1c")

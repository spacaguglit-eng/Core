# -*- coding: utf-8 -*-
"""
catalog.py — централизованный «Каталог» продуктов:
- нормализация наименований
- простейший парсинг (база/объём/бренд)
- скорость по продукту/линии с fallback на дефолт по линии
- просмотр парсинга (окно с таблицей)
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, Optional, Tuple
import re
import tkinter as tk
from tkinter import ttk, Toplevel

from product_parse import parse_product_name

__all__ = ["ProductKey", "Catalog", "make_default_catalog"]


# ======================= УТИЛИТЫ =============================================

def _norm_spaces(s: str) -> str:
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s.strip())
    return s


def _norm_quotes(s: str) -> str:
    s = s.replace("«", '"').replace("»", '"').replace("“", '"').replace("”", '"').replace("’", "'")
    s = s.replace("Ё", "Е").replace("ё", "е")
    return s


def _norm_dashes(s: str) -> str:
    return s.replace("–", "-").replace("—", "-")


def _cleanup(s: str) -> str:
    return _norm_spaces(_norm_quotes(_norm_dashes(str(s))))


# ======================= КЛЮЧ ПРОДУКТА =======================================

@dataclass(frozen=True)
class ProductKey:
    base: str
    volume: str
    brand: str

    @property
    def label(self) -> str:
        parts = [p for p in [self.base, self.volume, self.brand] if p]
        return " ".join(parts)


# ======================= ОСНОВНОЙ КЛАСС =====================================

class Catalog:
    """
    aliases: синоним → канон
    product_speeds: (канон, линия) → скорость
    line_defaults: линия → дефолтная скорость
    product_meta: (канон, линия) → {"container": str, "limit": Optional[float], "action": str}
    """
    def __init__(
        self,
        aliases: Optional[Dict[str, str]] = None,
        product_speeds: Optional[Dict[Tuple[str, str], float]] = None,
        line_defaults: Optional[Dict[str, float]] = None,
    ) -> None:
        self.aliases: Dict[str, str] = {self._canon(k): self._canon(v) for k, v in (aliases or {}).items()}
        self.product_speeds: Dict[Tuple[str, str], float] = dict(product_speeds or {})
        self.line_defaults: Dict[str, float] = {self._canon_line(k): v for k, v in (line_defaults or {}).items()}
        self.product_meta: Dict[Tuple[str, str], Dict[str, object]] = {}

    # ===== нормализация =====
    def normalize_name(self, name: str) -> str:
        n = _cleanup(name)
        n = self.aliases.get(self._canon(n), n)
        return n

    def _canon(self, s: str) -> str:
        return _cleanup(s).lower()

    def _canon_line(self, line: str) -> str:
        s = self._canon(line)
        m = re.search(r"(линия|line)\s*0*(\d+)", s)
        return f"линия {int(m.group(2))}" if m else s

    # ===== парсинг =====
    def parse_title(self, name: str) -> ProductKey:
        n = self.normalize_name(name)
        vol = ""
        m = re.search(r"(\d+(?:[.,]\d+)?)\s*(л|кг|г)\b", n, flags=re.IGNORECASE)
        if m:
            val, unit = m.group(1), m.group(2)
            vol = f"{val.replace('.', ',')} {unit}"
            n = _cleanup(n.replace(m.group(0), " "))

        brand = ""
        m2 = re.search(r'ТМ\s*["«](.*?)["»]', n, flags=re.IGNORECASE)
        if m2:
            brand = m2.group(1)
            n = _cleanup(n.replace(m2.group(0), " "))

        base = _cleanup(n)
        return ProductKey(base=base, volume=vol, brand=brand)

    # ===== скорость =====
    def speed(self, line: str, name: str) -> Optional[float]:
        ln = self._canon_line(line or "")
        nm = self.normalize_name(name or "")
        if (nm, ln) in self.product_speeds:
            return self.product_speeds[(nm, ln)]
        return self.line_defaults.get(ln)

    # ===== обновление =====
    def set_line_defaults(self, defaults: Dict[str, float]) -> None:
        self.line_defaults = {self._canon_line(k): v for k, v in (defaults or {}).items()}

    def set_product_speeds(self, product_speeds: Dict[Tuple[str, str], float]) -> None:
        self.product_speeds = dict(product_speeds or {})

    # ===== строки каталога для GUI ============================================
    def upsert(self, name: str, line: str, *, container: str = "", speed: Optional[float] = None,
               limit: Optional[float] = None, action: str = "") -> None:
        nm = self.normalize_name(name or "")
        ln = self._canon_line(line or "")
        if speed is not None:
            try:
                self.product_speeds[(nm, ln)] = float(speed)
            except Exception:
                pass
        meta = self.product_meta.get((nm, ln), {})
        meta.update({
            "container": str(container or ""),
            "limit": (float(limit) if (limit is not None and str(limit).strip() != "") else None),
            "action": str(action or "")
        })
        self.product_meta[(nm, ln)] = meta

    def rows(self):
        """ Экспорт строк для таблицы GUI. """
        out = []
        keys = set(self.product_speeds.keys()) | set(self.product_meta.keys())
        for k in sorted(keys):
            nm, ln = k
            meta = self.product_meta.get(k, {})
            out.append({
                "name": nm,
                "line": ln,
                "container": meta.get("container", ""),
                "speed": self.product_speeds.get(k, None),
                "limit": meta.get("limit", None),
                "action": meta.get("action", "")
            })
        return out

    def import_rows(self, rows: list[dict]) -> None:
        """ Полная замена содержимого из GUI. """
        self.product_speeds.clear()
        self.product_meta.clear()
        for r in rows:
            self.upsert(
                r.get("name", ""),
                r.get("line", ""),
                container=r.get("container", ""),
                speed=r.get("speed", None),
                limit=r.get("limit", None),
                action=r.get("action", ""),
            )

    def add_alias(self, src: str, dst: str) -> None:
        self.aliases[self._canon(src)] = self._canon(dst)

    # ===================== ПАРСИНГ GUI ========================================

    def show_parsing_window(self):
        """Открывает окно с результатами парсинга каталога."""
        parsed = []
        for row in self.rows():
            res = parse_product_name(row.get("name", ""), row.get("container", ""))
            parsed.append({
                "name": row.get("name", ""),
                "type": res.get("type", ""),
                "flavor": res.get("flavor", ""),
                "brand": res.get("brand", ""),
                "volume": res.get("volume", ""),
            })

        win = Toplevel()
        win.title("Результаты парсинга каталога")
        win.geometry("850x600")

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        cols = ("name", "type", "flavor", "brand", "volume")
        tv = ttk.Treeview(frame, columns=cols, show="headings")
        headers = {
            "name": "Название",
            "type": "Тип",
            "flavor": "Вкус",
            "brand": "Бренд",
            "volume": "Объём"
        }
        for c in cols:
            tv.heading(c, text=headers[c])
        tv.column("name", width=260)
        tv.column("type", width=90, anchor="center")
        tv.column("flavor", width=220)
        tv.column("brand", width=120)
        tv.column("volume", width=80, anchor="center")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        tv.pack(fill="both", expand=True)

        # Подсветка проблемных строк (без типа или бренда)
        tv.tag_configure("missing", background="#ffe6e6")

        for r in parsed:
            tag = ""
            if not r["type"] or not r["flavor"]:
                tag = "missing"
            tv.insert("", "end", values=(r["name"], r["type"], r["flavor"], r["brand"], r["volume"]), tags=(tag,))

        ttk.Label(win, text=f"Всего записей: {len(parsed)}").pack(pady=5)

    def add_parsing_button(self, parent_frame):
        """Добавляет кнопку «Парсинг» рядом с другими кнопками в GUI."""
        btn_parse = ttk.Button(parent_frame, text="🔍 Парсинг", command=self.show_parsing_window)
        btn_parse.pack(side="left", padx=5, pady=5)


# ======================= ФАБРИКА =============================================

def make_default_catalog() -> Catalog:
    return Catalog()

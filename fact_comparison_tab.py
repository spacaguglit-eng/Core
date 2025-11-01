# -*- coding: utf-8 -*-
"""
fact_comparison_tab.py — Вкладка «Факт/План» для отображения хронологии событий производства
---------------------------------------------------------------------------------------------------
• Загружает данные из JSON файла (OEE данные)
• Включает простои из записей
• Выстраивает хронологическую последовательность событий:
  - Начало производства
  - Простои (с указанием причины, категории, длительности)
  - Конец производства
• Отображает события в хронологическом порядке
"""

from __future__ import annotations
import os
import json
import re
import datetime as dt
from typing import List, Dict, Any, Optional, Tuple
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Импортируем HEADERS из json_import_tab для доступа к данным
try:
    from json_import_tab import HEADERS as JSON_IMPORT_HEADERS
except ImportError:
    JSON_IMPORT_HEADERS = [
        "Job ID","Продукт","Линия","День","Смена","Начало","Конец","Длит (мин)",
        "Σ простой (мин)","% простоя","Событий","План. простой (мин)",
        "EffMin (мин)","Ном. скорость (ш)","Потолок (шт)","Факт (шт)","OEE, %"
    ]

# ---------------------------------------------------------------------
_THIS_DIR = os.path.dirname(__file__)
_SCHEDULE_JSON = os.path.join(_THIS_DIR, "schedule_data.json")
_SETTINGS_PATH = os.path.join(_THIS_DIR, "settings_oee.json")

# Колонки таблицы хронологии событий
TIMELINE_COLS = (
    "time", "event_type", "job_id", "product", "line", "duration", 
    "reason", "kind", "quantity", "status"
)

TIMELINE_HEADERS = (
    "Время", "Тип события", "Job ID", "Продукт", "Линия", "Длительность (мин)", 
    "Причина/Описание", "Категория", "Кол-во", "Статус"
)

# Колонки таблицы сопоставления план/факт
COMPARISON_COLS = (
    "job_id", "product", "line", "plan_start", "plan_end", "fact_start", "fact_end",
    "time_deviation", "plan_qty", "fact_qty", "qty_deviation", "status", "note"
)

COMPARISON_HEADERS = (
    "Job ID", "Продукт", "Линия", "План (начало)", "План (конец)", 
    "Факт (начало)", "Факт (конец)", "Отклонение (время)", 
    "План (кол-во)", "Факт (кол-во)", "Отклонение (кол-во)", "Статус", "Примечание"
)


def _load_settings() -> dict:
    """Загрузка настроек"""
    try:
        if os.path.isfile(_SETTINGS_PATH):
            with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
                return d if isinstance(d, dict) else {}
    except Exception:
        pass
    return {}


def _save_settings(d: dict) -> None:
    """Сохранение настроек"""
    try:
        with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _load_schedule() -> List[Dict[str, Any]]:
    """Загрузка расписания из schedule_data.json"""
    try:
        if os.path.isfile(_SCHEDULE_JSON):
            with open(_SCHEDULE_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
                elif isinstance(data, dict) and "schedule" in data:
                    return data["schedule"]
    except Exception as e:
        print(f"[ERROR] Ошибка загрузки расписания: {e}")
    return []


def _get_fact_from_import_tab(nb: ttk.Notebook) -> List[Dict[str, Any]]:
    """Получение данных факта из вкладки Импорт JSON"""
    try:
        # Ищем вкладку "Импорт JSON"
        for tab_id in nb.tabs():
            if nb.tab(tab_id, "text") == "Импорт JSON":
                tab = nb.nametowidget(tab_id)
                # Рекурсивно ищем Treeview
                def find_treeview(widget):
                    if isinstance(widget, ttk.Treeview):
                        return widget
                    for child in widget.winfo_children():
                        result = find_treeview(child)
                        if result:
                            return result
                    return None
                
                tree = find_treeview(tab)
                if tree:
                    # Получаем данные из Treeview
                    fact_data = []
                    for item_id in tree.get_children():
                        values = tree.item(item_id, "values")
                        if values and len(values) >= len(JSON_IMPORT_HEADERS):
                            fact_item = {}
                            for i, header in enumerate(JSON_IMPORT_HEADERS):
                                fact_item[header] = values[i] if i < len(values) else ""
                            fact_data.append(fact_item)
                    return fact_data
    except Exception as e:
        print(f"[ERROR] Ошибка получения данных из импорта: {e}")
    return []


def _flatten_payload(payload: Any) -> List[Dict[str, Any]]:
    """Преобразование JSON данных в плоский список записей"""
    if isinstance(payload, list):
        return [r for r in payload if isinstance(r, dict)]
    if isinstance(payload, dict):
        if "data" in payload and isinstance(payload["data"], list):
            return payload["data"]
        for v in payload.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return v
    return []


def _load_fact_from_json(path: str) -> List[Dict[str, Any]]:
    """Загрузка факта из JSON файла с исходными данными (включая простои)"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return _flatten_payload(data)
    except Exception as e:
        print(f"[ERROR] Ошибка загрузки факта: {e}")
    return []


def _minutes_from_hhmm(beg: str, end: str) -> int:
    """Вычисляет длительность в минутах между временем начала и окончания"""
    try:
        def _to_minutes(t):
            if not t or ":" not in t:
                return None
            t = t.strip()
            parts = re.split(r"[:.]", t)
            if len(parts) < 2:
                return None
            hh = int(parts[0])
            mm = int(parts[1])
            return hh * 60 + mm
        
        a = _to_minutes(beg)
        b = _to_minutes(end)
        
        if a is None or b is None:
            return 0
        
        # Учитываем переход через полночь
        if b < a:
            b += 24 * 60
        
        return max(b - a, 0)
    except Exception:
        return 0


def _parse_datetime(dt_str: str, date_hint: Optional[str] = None) -> Optional[dt.datetime]:
    """
    Парсинг даты и времени из строки
    Если dt_str содержит только время (без даты), используется date_hint для даты
    """
    if not dt_str:
        return None
    
    dt_str = dt_str.strip()
    
    # Различные форматы с датой
    formats_with_date = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%d.%m %H:%M:%S",  # Формат без года: "02.11 06:07"
        "%d.%m %H:%M",      # Формат без года: "02.11 06:07"
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%S.%f",
    ]
    
    # Пробуем форматы с датой
    for fmt in formats_with_date:
        try:
            parsed = dt.datetime.strptime(dt_str, fmt)
            # Если формат без года (например, "%d.%m %H:%M"), добавляем год из подсказки или текущий год
            if fmt in ("%d.%m %H:%M:%S", "%d.%m %H:%M"):
                if date_hint:
                    hint_date = _parse_datetime(date_hint)
                    if hint_date:
                        parsed = parsed.replace(year=hint_date.year)
                    else:
                        parsed = parsed.replace(year=dt.date.today().year)
                else:
                    parsed = parsed.replace(year=dt.date.today().year)
            return parsed
        except ValueError:
            continue
    
    # Если не получилось, пробуем только время (HH:MM или HH:MM:SS)
    if ":" in dt_str and len(dt_str) <= 8 and not any(c.isalpha() for c in dt_str):
        try:
            # Парсим время
            if dt_str.count(":") == 1:  # HH:MM
                time_obj = dt.datetime.strptime(dt_str, "%H:%M").time()
            else:  # HH:MM:SS
                time_obj = dt.datetime.strptime(dt_str, "%H:%M:%S").time()
            
            # Если есть подсказка с датой, используем её
            if date_hint:
                date_obj = _parse_datetime(date_hint)
                if date_obj:
                    return dt.datetime.combine(date_obj.date(), time_obj)
            
            # Если нет подсказки, используем текущую дату (для сравнения)
            # Но это не очень правильно - лучше использовать дату из плана
            return dt.datetime.combine(dt.date.today(), time_obj)
        except ValueError:
            pass
    
    return None


def _get_time_sort_key(time_str: str) -> Tuple[int, int, int]:
    """Извлекает время для сортировки в формате (часы, минуты, секунды)"""
    if not time_str or ":" not in time_str:
        return (0, 0, 0)
    
    try:
        parts = re.split(r"[:.]", time_str.strip())
        hh = int(parts[0]) if len(parts) > 0 else 0
        mm = int(parts[1]) if len(parts) > 1 else 0
        ss = int(parts[2]) if len(parts) > 2 else 0
        return (hh, mm, ss)
    except Exception:
        return (0, 0, 0)


def _build_timeline_events(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Строит хронологию событий из записей JSON (включая простои)
    Возвращает список событий, отсортированный по времени
    """
    events = []
    
    for record in records:
        job_id = str(record.get("job_id", "")).strip()
        product = str(record.get("product", "")).strip()
        line = str(record.get("line", "")).strip()
        start = str(record.get("start", "")).strip()
        end = str(record.get("end", "")).strip()
        quantity = record.get("fact", record.get("fact_qty", record.get("quantity", 0)))
        
        # Начало производства
        if start:
            events.append({
                "time": start,
                "event_type": "▶ Начало производства",
                "job_id": job_id,
                "product": product,
                "line": line,
                "duration": "",
                "reason": "",
                "kind": "",
                "quantity": quantity,
                "status": "🟢",
            })
        
        # Простои
        downtimes = record.get("downtimes", [])
        if isinstance(downtimes, list):
            for dt_item in downtimes:
                if isinstance(dt_item, dict):
                    dt_start = str(dt_item.get("start", dt_item.get("beg", ""))).strip()
                    dt_end = str(dt_item.get("end", dt_item.get("stop", ""))).strip()
                    dt_reason = str(dt_item.get("reason", "")).strip()
                    dt_kind = str(dt_item.get("kind", dt_item.get("type", dt_item.get("category", "")))).strip()
                    dt_desc = str(dt_item.get("description", dt_item.get("desc", dt_item.get("comment", "")))).strip()
                    
                    # Длительность простоя
                    dt_duration = dt_item.get("duration_min", dt_item.get("duration", 0))
                    if not dt_duration and dt_start and dt_end:
                        dt_duration = _minutes_from_hhmm(dt_start, dt_end)
                    
                    # Описание простоя
                    dt_display = dt_reason if dt_reason else dt_desc
                    if not dt_display:
                        dt_display = "Простой"
                    
                    # Используем время начала простоя
                    if dt_start:
                        events.append({
                            "time": dt_start,
                            "event_type": "⏸ Простой",
                            "job_id": job_id,
                            "product": product,
                            "line": line,
                            "duration": f"{int(dt_duration)}" if dt_duration else "",
                            "reason": dt_display,
                            "kind": dt_kind,
                            "quantity": "",
                            "status": "🟡",
                        })
        
        # Конец производства
        if end:
            events.append({
                "time": end,
                "event_type": "■ Конец производства",
                "job_id": job_id,
                "product": product,
                "line": line,
                "duration": "",
                "reason": "",
                "kind": "",
                "quantity": quantity,
                "status": "🔴",
            })
    
    # Сортируем события по времени
    events.sort(key=lambda e: _get_time_sort_key(e.get("time", "")))
    
    return events


def _calculate_time_deviation(plan_start: Optional[str], plan_end: Optional[str],
                             fact_start: Optional[str], fact_end: Optional[str]) -> Tuple[Optional[int], str]:
    """
    Расчет отклонения времени в минутах
    Возвращает: (отклонение_в_минутах, описание)
    """
    # Парсим плановое время
    plan_start_dt = _parse_datetime(plan_start) if plan_start else None
    plan_end_dt = _parse_datetime(plan_end) if plan_end else None
    
    # Для фактического времени используем дату из плана, если факт содержит только время
    fact_start_dt = None
    fact_end_dt = None
    
    if fact_start:
        # Если факт содержит только время (без даты), используем дату из плана
        fact_start_dt = _parse_datetime(fact_start, plan_start)
    
    if fact_end:
        # Если факт содержит только время (без даты), используем дату из плана
        fact_end_dt = _parse_datetime(fact_end, plan_end)
    
    if not plan_start_dt or not plan_end_dt:
        return None, "Нет плана"
    
    if not fact_start_dt or not fact_end_dt:
        return None, "Нет факта"
    
    # Плановая длительность
    plan_duration = (plan_end_dt - plan_start_dt).total_seconds() / 60
    
    # Фактическая длительность
    fact_duration = (fact_end_dt - fact_start_dt).total_seconds() / 60
    
    # Отклонение по началу
    start_deviation = (fact_start_dt - plan_start_dt).total_seconds() / 60
    
    # Отклонение по длительности
    duration_deviation = fact_duration - plan_duration
    
    deviation_minutes = int(start_deviation)
    
    if abs(deviation_minutes) < 5 and abs(duration_deviation) < 5:
        status = "OK"
    elif deviation_minutes > 0:
        status = f"Задержка {deviation_minutes:.0f} мин"
    else:
        status = f"Опережение {abs(deviation_minutes):.0f} мин"
    
    return deviation_minutes, status


def _calculate_qty_deviation(plan_qty: Optional[float], fact_qty: Optional[float]) -> Tuple[Optional[float], str]:
    """
    Расчет отклонения количества
    Возвращает: (отклонение, описание)
    """
    if plan_qty is None or plan_qty == 0:
        return None, "Нет плана"
    
    if fact_qty is None:
        return None, "Нет факта"
    
    deviation = fact_qty - plan_qty
    percent = (deviation / plan_qty) * 100 if plan_qty > 0 else 0
    
    if abs(percent) < 1:
        status = "OK"
    elif percent > 0:
        status = f"+{percent:.1f}%"
    else:
        status = f"{percent:.1f}%"
    
    return deviation, status


def _normalize_job_id(job_id: str) -> Tuple[str, str]:
    """
    Нормализация job_id для сравнения
    Возвращает: (базовый_id, суффикс)
    Например: "JOB-001-P1" -> ("JOB-001", "-P1")
    """
    if not job_id:
        return "", ""
    
    job_id = str(job_id).strip()
    
    # Убираем суффиксы типа -P1, -P2 (части работы)
    if "-P" in job_id.upper() or "-PART" in job_id.upper():
        parts = job_id.rsplit("-", 1)
        if len(parts) == 2 and parts[1][0].upper() == "P":
            base_id = parts[0]
            suffix = "-" + parts[1]
            return base_id, suffix
    
    return job_id, ""


def _match_schedule_with_fact(schedule: List[Dict], fact: List[Dict]) -> List[Dict[str, Any]]:
    """
    Сопоставление расписания с фактом по job_id
    Возвращает список сравнений
    """
    results = []
    
    # Создаем индекс факта по job_id (точное совпадение)
    fact_index_exact = {}
    # Создаем индекс факта по базовому job_id (для частей работы)
    fact_index_base = {}
    
    for fact_item in fact:
        job_id_str = str(fact_item.get("Job ID", fact_item.get("job_id", ""))).strip()
        if job_id_str:
            # Точное совпадение
            fact_index_exact[job_id_str] = fact_item
            
            # Базовое совпадение (без суффиксов)
            base_id, suffix = _normalize_job_id(job_id_str)
            if base_id and base_id not in fact_index_base:
                fact_index_base[base_id] = fact_item
    
    # Проходим по расписанию
    for plan_item in schedule:
        job_id = str(plan_item.get("job_id", "")).strip()
        if not job_id:
            continue
        
        # Ищем факт: сначала точное совпадение, потом по базовому ID
        fact_item = fact_index_exact.get(job_id)
        
        if not fact_item:
            # Пробуем найти по базовому ID (для частей работы типа JOB-001-P1)
            base_id, suffix = _normalize_job_id(job_id)
            if base_id:
                fact_item = fact_index_base.get(base_id)
        
        # План
        plan_start = plan_item.get("start", "")
        plan_end = plan_item.get("end", "")
        plan_qty = plan_item.get("qty", "")
        try:
            plan_qty_num = float(plan_qty) if plan_qty else None
        except (ValueError, TypeError):
            plan_qty_num = None
        
        # Факт
        if fact_item:
            # Пробуем разные варианты имен полей для времени
            fact_start = (fact_item.get("Начало") or fact_item.get("start") or 
                         fact_item.get("Start") or fact_item.get("begin") or "")
            fact_end = (fact_item.get("Конец") or fact_item.get("end") or 
                       fact_item.get("End") or fact_item.get("finish") or "")
            
            # Пробуем разные варианты имен полей для количества
            fact_qty = (fact_item.get("Факт (шт)") or fact_item.get("fact_qty") or 
                       fact_item.get("fact") or fact_item.get("qty") or 
                       fact_item.get("quantity") or fact_item.get("Факт") or "")
            
            try:
                fact_qty_num = float(fact_qty) if fact_qty else None
            except (ValueError, TypeError):
                fact_qty_num = None
        else:
            fact_start = ""
            fact_end = ""
            fact_qty_num = None
        
        # Расчет отклонений по времени
        time_deviation, time_status = _calculate_time_deviation(
            plan_start, plan_end, fact_start, fact_end
        )
        
        # Отладка: проверяем что получилось
        # print(f"DEBUG: job_id={job_id}, plan_start={plan_start}, fact_start={fact_start}, time_deviation={time_deviation}, time_status={time_status}")
        
        qty_deviation, qty_status = _calculate_qty_deviation(plan_qty_num, fact_qty_num)
        
        # Общий статус
        if not fact_item:
            overall_status = "❌ Нет факта"
        elif time_status == "OK" and qty_status == "OK":
            overall_status = "✅ OK"
        elif time_status != "OK" and qty_status != "OK":
            overall_status = "⚠️ Отклонение"
        elif time_status != "OK":
            overall_status = "⚠️ Время"
        else:
            overall_status = "⚠️ Количество"
        
        # Примечание
        note_parts = []
        if not fact_item:
            note_parts.append("Не найден в факте")
        else:
            # Проверяем, использовали ли базовый ID для сопоставления
            base_id, suffix = _normalize_job_id(job_id)
            if suffix and fact_item.get("Job ID", "") != job_id:
                fact_job_id = fact_item.get("Job ID", "")
                note_parts.append(f"Совпадение по базовому ID: {base_id}")
            
            if time_status != "OK" and time_status not in ("Нет плана", "Нет факта"):
                note_parts.append(f"Время: {time_status}")
            if qty_status != "OK" and qty_status not in ("Нет плана", "Нет факта"):
                note_parts.append(f"Количество: {qty_status}")
        
        note = "; ".join(note_parts) if note_parts else ""
        
        results.append({
            "job_id": job_id,
            "product": plan_item.get("name", ""),
            "line": plan_item.get("line", ""),
            "plan_start": plan_start,
            "plan_end": plan_end,
            "fact_start": fact_start,
            "fact_end": fact_end,
            "time_deviation": (
                f"{int(time_deviation)} мин" if time_deviation is not None and time_status != "OK" 
                else ("OK" if time_status == "OK" 
                      else time_status if time_status in ("Нет плана", "Нет факта") 
                      else "")
            ),
            "plan_qty": f"{plan_qty_num:.0f}" if plan_qty_num is not None else "",
            "fact_qty": f"{fact_qty_num:.0f}" if fact_qty_num is not None else "",
            "qty_deviation": f"{qty_deviation:.0f}" if qty_deviation is not None else "",
            "status": overall_status,
            "note": note,
        })
    
    # Также добавляем записи факта, которые не были найдены в расписании
    used_fact_job_ids = set()
    for result in results:
        job_id = result.get("job_id", "")
        if job_id:
            used_fact_job_ids.add(job_id)
            # Также добавляем базовый ID, если был суффикс
            base_id, suffix = _normalize_job_id(job_id)
            if base_id:
                used_fact_job_ids.add(base_id)
    
    for fact_item in fact:
        fact_job_id = str(fact_item.get("Job ID", fact_item.get("job_id", ""))).strip()
        if fact_job_id and fact_job_id not in used_fact_job_ids:
            # Факт есть, но плана нет
            fact_start = (fact_item.get("Начало") or fact_item.get("start") or 
                         fact_item.get("Start") or fact_item.get("begin") or "")
            fact_end = (fact_item.get("Конец") or fact_item.get("end") or 
                       fact_item.get("End") or fact_item.get("finish") or "")
            fact_qty = (fact_item.get("Факт (шт)") or fact_item.get("fact_qty") or 
                       fact_item.get("fact") or fact_item.get("qty") or 
                       fact_item.get("quantity") or fact_item.get("Факт") or "")
            
            try:
                fact_qty_num = float(fact_qty) if fact_qty else None
            except (ValueError, TypeError):
                fact_qty_num = None
            
            results.append({
                "job_id": fact_job_id,
                "product": fact_item.get("Продукт", fact_item.get("product", "")),
                "line": fact_item.get("Линия", fact_item.get("line", "")),
                "plan_start": "",
                "plan_end": "",
                "fact_start": fact_start,
                "fact_end": fact_end,
                "time_deviation": "Нет плана",
                "plan_qty": "",
                "fact_qty": f"{fact_qty_num:.0f}" if fact_qty_num is not None else "",
                "qty_deviation": "",
                "status": "❌ Нет плана",
                "note": "Запись найдена только в факте",
            })
    
    return results


class FactComparisonTab:
    """Вкладка для сравнения факта с планом"""
    
    def __init__(self, parent: ttk.Frame, parent_notebook: Optional[ttk.Notebook] = None):
        self.parent = parent
        self.parent_notebook = parent_notebook
        self.fact_json_path: Optional[str] = None
        
        # Загружаем настройки
        settings = _load_settings()
        self.fact_json_path = settings.get("oee_json_path", "")
        
        self._build_ui()
        self._refresh_comparison()
    
    def _build_ui(self):
        """Построение интерфейса"""
        # Верхняя панель управления
        toolbar = ttk.Frame(self.parent)
        toolbar.pack(fill="x", padx=8, pady=(8, 4))
        
        # Переключатель режимов
        mode_frame = ttk.LabelFrame(toolbar, text="Режим отображения", padding=6)
        mode_frame.pack(side="left", padx=(0, 12))
        
        self.view_mode = tk.StringVar(value="timeline")
        ttk.Radiobutton(mode_frame, text="📅 Хронология", variable=self.view_mode, 
                       value="timeline", command=self._on_mode_change).pack(side="left", padx=(0, 8))
        ttk.Radiobutton(mode_frame, text="⚖️ Сопоставление", variable=self.view_mode, 
                       value="comparison", command=self._on_mode_change).pack(side="left")
        
        ttk.Button(toolbar, text="🔄 Обновить", 
                   command=self._refresh_comparison).pack(side="left", padx=(0, 8))
        
        self.lbl_info = ttk.Label(toolbar, text="Хронология событий из JSON импорта", foreground="#666")
        self.lbl_info.pack(side="left", padx=(0, 8))
        
        # Таблица сравнения
        table_frame = ttk.Frame(self.parent)
        table_frame.pack(fill="both", expand=True, padx=8, pady=4)
        
        # Treeview
        self.tree = ttk.Treeview(table_frame, show="headings", height=20)
        
        # Scrollbars
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Размещение
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        
        # Настройка колонок будет меняться в зависимости от режима
        # Теги для цветовой индикации событий
        self.tree.tag_configure("production_start", background="#e8f5e9", foreground="#2e7d32")
        self.tree.tag_configure("downtime", background="#fff3e0", foreground="#e65100")
        self.tree.tag_configure("production_end", background="#ffebee", foreground="#c62828")
        
        # Теги для сопоставления план/факт
        self.tree.tag_configure("ok", background="#e8f5e9")
        self.tree.tag_configure("warning", background="#fff3e0")
        self.tree.tag_configure("error", background="#ffebee")
        self.tree.tag_configure("no_fact", background="#f5f5f5")
        
        # Инициализируем колонки для режима по умолчанию
        self._setup_timeline_columns()
    
    def _setup_timeline_columns(self):
        """Настройка колонок для режима хронологии"""
        self.tree["columns"] = TIMELINE_COLS
        column_widths = {
            "time": 150,
            "event_type": 180,
            "job_id": 120,
            "product": 280,
            "line": 100,
            "duration": 120,
            "reason": 250,
            "kind": 150,
            "quantity": 100,
            "status": 80,
        }
        
        for col, header in zip(TIMELINE_COLS, TIMELINE_HEADERS):
            self.tree.heading(col, text=header)
            anchor = "center" if col in ("status", "line", "duration", "quantity") else "w"
            self.tree.column(col, width=column_widths.get(col, 120), anchor=anchor)
    
    def _setup_comparison_columns(self):
        """Настройка колонок для режима сопоставления"""
        self.tree["columns"] = COMPARISON_COLS
        column_widths = {
            "job_id": 100,
            "product": 250,
            "line": 100,
            "plan_start": 150,
            "plan_end": 150,
            "fact_start": 150,
            "fact_end": 150,
            "time_deviation": 180,  # Увеличена ширина для отображения отклонений
            "plan_qty": 100,
            "fact_qty": 100,
            "qty_deviation": 120,
            "status": 100,
            "note": 200,
        }
        
        for col, header in zip(COMPARISON_COLS, COMPARISON_HEADERS):
            self.tree.heading(col, text=header)
            # Для отклонений выравнивание по центру
            anchor = "center" if col in ("status", "job_id", "line", "time_deviation", "qty_deviation") else "w"
            self.tree.column(col, width=column_widths.get(col, 120), anchor=anchor, minwidth=80)
    
    def _on_mode_change(self):
        """Переключение режима отображения"""
        self._refresh_comparison()
    
    def _refresh_comparison(self):
        """Обновление данных в зависимости от режима"""
        # Очищаем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        mode = self.view_mode.get()
        
        if mode == "timeline":
            self._refresh_timeline()
        else:
            self._refresh_comparison_mode()
    
    def _refresh_timeline(self):
        """Обновление хронологии событий"""
        # Настраиваем колонки для хронологии
        self._setup_timeline_columns()
        
        # Получаем исходные данные из JSON файла
        records = []
        
        # Сначала пробуем загрузить из файла (основной источник)
        if self.fact_json_path and os.path.isfile(self.fact_json_path):
            try:
                records = _load_fact_from_json(self.fact_json_path)
            except Exception as e:
                self.lbl_info.config(text=f"Ошибка загрузки JSON: {e}", foreground="#d32f2f")
                return
        
        if not records:
            self.lbl_info.config(
                text="Данные не загружены. Загрузите JSON файл во вкладке 'Импорт JSON'.", 
                foreground="#f57c00"
            )
            return
        
        # Строим хронологию событий
        events = _build_timeline_events(records)
        
        if not events:
            self.lbl_info.config(text="События не найдены в данных", foreground="#f57c00")
            return
        
        # Добавляем события в таблицу
        for event in events:
            # Определяем тег по типу события
            event_type = event.get("event_type", "")
            if "Начало" in event_type:
                tag = "production_start"
            elif "Простой" in event_type:
                tag = "downtime"
            elif "Конец" in event_type:
                tag = "production_end"
            else:
                tag = ""
            
            # Форматируем количество
            qty = event.get("quantity", "")
            if qty and isinstance(qty, (int, float)):
                qty = str(int(qty)) if isinstance(qty, float) and qty.is_integer() else str(qty)
            
            values = [
                event.get("time", ""),
                event.get("event_type", ""),
                event.get("job_id", ""),
                event.get("product", ""),
                event.get("line", ""),
                event.get("duration", ""),
                event.get("reason", ""),
                event.get("kind", ""),
                qty if qty else "",
                event.get("status", ""),
            ]
            
            self.tree.insert("", "end", values=values, tags=(tag,))
        
        # Статистика
        total_events = len(events)
        start_count = sum(1 for e in events if "Начало" in e.get("event_type", ""))
        downtime_count = sum(1 for e in events if "Простой" in e.get("event_type", ""))
        end_count = sum(1 for e in events if "Конец" in e.get("event_type", ""))
        
        self.lbl_info.config(
            text=f"Хронология: Всего событий {total_events} | Начал: {start_count} | Простоев: {downtime_count} | Завершений: {end_count}",
            foreground="#388e3c"
        )
    
    def _refresh_comparison_mode(self):
        """Обновление сопоставления план/факт"""
        # Настраиваем колонки для сопоставления
        self._setup_comparison_columns()
        
        # Загружаем расписание
        schedule = _load_schedule()
        if not schedule:
            self.lbl_info.config(text="Расписание не найдено. Сначала создайте расписание.", foreground="#d32f2f")
            return
        
        # Получаем факт из JSON файла
        fact = []
        if self.fact_json_path and os.path.isfile(self.fact_json_path):
            try:
                fact = _load_fact_from_json(self.fact_json_path)
            except Exception as e:
                self.lbl_info.config(text=f"Ошибка загрузки JSON: {e}", foreground="#d32f2f")
                return
        
        # Если не получили из файла - пробуем из вкладки Импорт JSON
        if not fact and self.parent_notebook:
            fact = _get_fact_from_import_tab(self.parent_notebook)
        
        if not fact:
            self.lbl_info.config(
                text="Факт не загружен. Загрузите данные во вкладке 'Импорт JSON'.", 
                foreground="#f57c00"
            )
            return
        
        # Сопоставляем план с фактом
        comparisons = _match_schedule_with_fact(schedule, fact)
        
        # Добавляем в таблицу
        for comp in comparisons:
            # Определяем тег для цвета
            status = comp.get("status", "")
            if "Нет факта" in status:
                tag = "no_fact"
            elif "Нет плана" in status:
                tag = "error"
            elif "OK" in status:
                tag = "ok"
            elif "⚠️" in status:
                tag = "warning"
            else:
                tag = "error"
            
            values = [comp.get(col, "") for col in COMPARISON_COLS]
            self.tree.insert("", "end", values=values, tags=(tag,))
        
        # Статистика
        total = len(comparisons)
        ok_count = sum(1 for c in comparisons if "OK" in c.get("status", ""))
        warning_count = sum(1 for c in comparisons if "⚠️" in c.get("status", ""))
        no_fact_count = sum(1 for c in comparisons if "Нет факта" in c.get("status", ""))
        no_plan_count = sum(1 for c in comparisons if "Нет плана" in c.get("status", ""))
        
        self.lbl_info.config(
            text=f"Сопоставление: Всего: {total} | ✅ OK: {ok_count} | ⚠️ Отклонения: {warning_count} | ❌ Нет факта: {no_fact_count} | ❌ Нет плана: {no_plan_count}",
            foreground="#388e3c" if ok_count == total else "#f57c00" if warning_count > 0 else "#666"
        )


def show_fact_comparison_tab(parent_notebook: ttk.Notebook):
    """Создание вкладки сравнения факта с планом в planning_tab"""
    # Находим вкладку "Планирование"
    planning_tab = None
    for tid in parent_notebook.tabs():
        if parent_notebook.tab(tid, "text") == "Планирование":
            planning_tab = parent_notebook.nametowidget(tid)
            break
    
    if not planning_tab:
        return
    
    # Находим подвкладки (sub notebook)
    for child in planning_tab.winfo_children():
        if isinstance(child, ttk.Notebook):
            # Добавляем новую вкладку "Факт/План"
            tab_fact = ttk.Frame(child)
            child.add(tab_fact, text="Факт/План")
            
            # Создаем экземпляр вкладки сравнения
            try:
                FactComparisonTab(tab_fact)
            except Exception as e:
                import traceback
                traceback.print_exc()
                ttk.Label(tab_fact, text=f"Ошибка при инициализации сравнения: {e}", foreground="#a00")\
                   .pack(anchor="w", padx=8, pady=8)
            
            break


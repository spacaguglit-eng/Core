# -*- coding: utf-8 -*-
"""
planning_tab.py — вкладка «Планирование» (План / Расписание / Импорт)

Главное:
- План: редактирование, сортировка, сохранение/загрузка JSON
- Импорт: распознавание нескольких паттернов
    • Excel-TSV/CSV (в т.ч. «Имя<TAB/много пробелов>Кол-во»)
    • Письмо/чистый текст (в т.ч. «Сироп со вкусом и ароматом "Ваниль" …»)
    • Строки «CIP 1/2» игнорируются
- Автонормализация объёма («0,25» → «0,25 л»), чисел, брендов/вкусов
- Сопоставление с каталогом (catalog_data.json / catalog.json)
- Импорт в План с полями Type / Flavor / Brand
"""

from __future__ import annotations
import os, re, json
from typing import List, Dict, Any, Tuple, Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
# мягкое подключение парсера продукта
try:
    from product_parse import parse_product_name as _pparse  # type: ignore
except Exception:
    _pparse = None

# пути
_THIS_DIR = os.path.dirname(__file__)
_PLAN_JSON = os.path.join(_THIS_DIR, "jobs_plan.json")
_CATALOG_JSON_MAIN = os.path.join(_THIS_DIR, "catalog_data.json")
_CATALOG_JSON_FALL = os.path.join(_THIS_DIR, "catalog.json")

def _catalog_path() -> str:
    return _CATALOG_JSON_MAIN if os.path.isfile(_CATALOG_JSON_MAIN) else _CATALOG_JSON_FALL

# ====== колонки ПЛАНА ========================================================
COL_KEYS: Tuple[str, ...] = (
    "priority","job_id","name",
    "volume","flavor","brand","type",
    "quantity","line",
    "speed","speed_source",
    "status","fact_qty","progress",
)
COL_HEADERS: Tuple[str, ...] = (
    "Приоритет","ID задания","Наименование",
    "Объём","Вкус","Бренд","Тип",
    "Кол-во","Линия",
    "Скорость","Источник",
    "Статус","Факт, шт","Прогресс",
)
COL_WIDTHS: Tuple[int, ...] = (
    80,
    120, 340,
    100, 240, 140, 120,
    90, 120,
    90, 110,
    90, 90, 120,
)
_NUMERIC_COLS = {"quantity","speed","fact_qty","priority"}

# ====== сортировка Treeview ==================================================
_SORT_STATE: dict[tuple[int, str], bool] = {}
def _nat_key(s: str):
    parts = re.findall(r"\d+|\D+", str(s))
    out = []
    for p in parts:
        out.append((0,int(p)) if p.isdigit() else (1,p.lower()))
    return tuple(out)

def _enable_tree_sort(tree: ttk.Treeview):
    """ОТКЛЮЧЕНО - сортировка нарушает порядок! Используйте drag & drop"""
    # Сортировка отключена чтобы сохранять порядок из файла
    pass

# ====== Drag & Drop для перетаскивания строк ==================================
def _enable_drag_and_drop(tree: ttk.Treeview, on_reorder_callback=None):
    """
    Включает Drag & Drop для перетаскивания строк в Treeview
    
    Args:
        tree: Treeview виджет
        on_reorder_callback: функция, вызываемая после изменения порядка
    """
    drag_data = {"item": None, "y": 0, "start_time": 0, "moved": False}
    
    def on_drag_start(event):
        """Начало перетаскивания"""
        import time
        item = tree.identify_row(event.y)
        if item:
            drag_data["item"] = item
            drag_data["y"] = event.y
            drag_data["start_time"] = time.time()
            drag_data["moved"] = False
            # Подсвечиваем перетаскиваемую строку
            tree.selection_set(item)
    
    def on_drag_motion(event):
        """Перемещение во время перетаскивания"""
        if not drag_data["item"]:
            return
        
        # Определяем было ли движение
        if abs(event.y - drag_data["y"]) > 5:  # минимальное расстояние для drag
            drag_data["moved"] = True
        
        # Определяем над какой строкой сейчас курсор
        target_item = tree.identify_row(event.y)
        if target_item and target_item != drag_data["item"]:
            # Показываем куда будет вставлена строка
            tree.selection_set(target_item)
    
    def on_drag_release(event):
        """Завершение перетаскивания"""
        import time
        
        if not drag_data["item"]:
            return
        
        # Если не было движения или прошло мало времени - это клик, не drag
        elapsed = time.time() - drag_data["start_time"]
        if not drag_data["moved"] or elapsed < 0.1:
            drag_data["item"] = None
            return
        
        source_item = drag_data["item"]
        target_item = tree.identify_row(event.y)
        
        if target_item and source_item != target_item:
            # Определяем позицию для вставки
            source_parent = tree.parent(source_item)
            target_parent = tree.parent(target_item)
            
            # Проверяем что оба элемента в одном уровне (не группа/запись)
            source_is_group = tree.item(source_item, "text").startswith("📍")
            target_is_group = tree.item(target_item, "text").startswith("📍")
            
            if source_is_group == target_is_group and source_parent == target_parent:
                # Получаем индекс целевой строки
                all_items = list(tree.get_children(target_parent if target_parent else ""))
                target_index = all_items.index(target_item)
                
                # Если перетаскиваем вниз, вставляем ПОСЛЕ целевой строки
                if event.y > drag_data["y"]:
                    target_index += 1
                
                # Перемещаем строку
                tree.move(source_item, target_parent if target_parent else "", target_index)
                tree.selection_set(source_item)
                
                # Вызываем callback если есть
                if on_reorder_callback:
                    on_reorder_callback()
        
        # Сбрасываем данные перетаскивания
        drag_data["item"] = None
        drag_data["y"] = 0
        drag_data["moved"] = False
    
    # Привязываем события (добавляем "+", чтобы не перезаписывать существующие обработчики)
    tree.bind("<Button-1>", on_drag_start, add="+")
    tree.bind("<B1-Motion>", on_drag_motion, add="+")
    tree.bind("<ButtonRelease-1>", on_drag_release, add="+")

def _autofit_columns(tree: ttk.Treeview):
    """Автоматическая подгонка ширины столбцов"""
    for col in tree["columns"]:
        max_width = 0
        # Проверяем заголовок
        hdr_width = len(tree.heading(col).get("text", col))
        max_width = max(max_width, hdr_width)
        
        # Проверяем данные
        for parent in tree.get_children(""):
            if tree.item(parent, "text").startswith("📍"):
                # Обходим все записи внутри групп
                for item in tree.get_children(parent):
                    value = str(tree.set(item, col))
                    max_width = max(max_width, len(value))
            else:
                # Прямые записи
                value = str(tree.set(parent, col))
                max_width = max(max_width, len(value))
        
        # Вычисляем оптимальную ширину
        calculated_width = max(max_width * 8 + 30, 60)
        # Ограничиваем диапазон
        final_width = min(max(calculated_width, 60), 500)
        tree.column(col, width=final_width)

def _config_tree(tree: ttk.Treeview, cols, headers, widths, numeric_cols):
    tree.configure(columns=cols, show="headings", selectmode="extended")
    for key, hdr, w in zip(cols, headers, widths):
        tree.heading(key, text=hdr)
        tree.column(key, width=w, anchor=("e" if key in numeric_cols else "w"))
    _enable_tree_sort(tree)
def _norm_line_to_num(line: str) -> int:
    """Из строки 'Линия 3' / '3' / 'L3' вытащить номер линии, иначе 0."""
    m = re.search(r'(\d+)', str(line or ""))
    return int(m.group(1)) if m else 0

def _collect_existing_job_ids(tree: ttk.Treeview) -> set[str]:
    """Собрать уже используемые JobID из таблицы Плана."""
    ids = set()
    if "job_id" in COL_KEYS:
        jx = COL_KEYS.index("job_id")
        for iid in tree.get_children(""):
            vals = tree.item(iid, "values")
            if jx < len(vals) and vals[jx]:
                ids.add(str(vals[jx]))
    return ids

def _next_job_id(existing: set[str], line: str) -> str:
    """Сгенерировать уникальный JobID формата J-YYMMDD-LNN-XXX."""
    today = datetime.now().strftime("%y%m%d")
    ln = _norm_line_to_num(line)
    base = f"J-{today}-L{ln:02d}-"
    n = 1
    while True:
        jid = f"{base}{n:03d}"
        if jid not in existing:
            existing.add(jid)
            return jid
        n += 1

# ====== UI для импорта: два окна =============================================
def _create_import_panes(tab_import: ttk.Frame, top: ttk.Frame):
    split = ttk.Panedwindow(tab_import, orient="horizontal")
    split.pack(fill="both", expand=True, padx=8, pady=(0,8))
    left = ttk.Frame(split); split.add(left, weight=1)
    right= ttk.Frame(split); split.add(right, weight=2)

    ttk.Label(left, text="Вставьте текст / из Excel", foreground="#666").grid(
        row=0,column=0, sticky="w", padx=2, pady=(2,4))
    txt = tk.Text(left, wrap="word", height=10, undo=True)
    scL = ttk.Scrollbar(left, orient="vertical", command=txt.yview)
    txt.configure(yscrollcommand=scL.set)
    txt.grid(row=1, column=0, sticky="nsew"); scL.grid(row=1,column=1,sticky="ns")
    left.rowconfigure(1, weight=1); left.columnconfigure(0, weight=1)

    ttk.Label(right, text="Результаты распознавания", foreground="#666").grid(
        row=0,column=0, sticky="w", padx=2, pady=(2,4))
    tree = ttk.Treeview(right, show="headings", selectmode="extended")
    scY = ttk.Scrollbar(right, orient="vertical", command=tree.yview)
    scX = ttk.Scrollbar(right, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=scY.set, xscrollcommand=scX.set)
    tree.grid(row=1, column=0, sticky="nsew"); scY.grid(row=1,column=1,sticky="ns")
    scX.grid(row=2, column=0, sticky="ew")
    right.rowconfigure(1, weight=1); right.columnconfigure(0, weight=1)

    def _reset():
        tab_import.update_idletasks()
        w = split.winfo_width() or 900
        split.sashpos(0, int(w*0.42))
    split.bind("<Configure>", lambda _e: _reset())
    ttk.Button(top, text="↹ Сбросить окна", command=_reset).pack(side="left", padx=6)
    _reset()
    return txt, tree, split

# ====== утилиты нормализации ==================================================
def _coerce_numeric(val: str) -> str:
    if val is None: return ""
    s = str(val).strip().replace("\xa0","")
    s = s.replace(" ","")
    return s if s.isdigit() else (re.sub(r"[^\d]","",s) if re.search(r"\d",s) else "")

def _normalize_volume_to_str(vol: str | float | int) -> str:
    """
    Нормализация объёма/массы:
      • понимает л/кг/ml/мл
      • сохраняет исходную единицу (л или кг)
      • форматирует как '1,0 л' / '1,0 кг'
    """
    if vol is None:
        return ""
    s = str(vol).strip().replace("\xa0", " ").lower()

    # ищем число + (л|l|кг|kg|мл|ml) — unit может отсутствовать
    m = re.search(r"(\d+(?:[.,]\d+)?)(?:\s*(л|l|кг|kg|мл|ml))?$", s)
    if not m:
        return s

    num_raw = (m.group(1) or "").replace(" ", "").replace(",", ".")
    unit_raw = (m.group(2) or "").strip()

    # нормализуем юнит до «л» или «кг» (ml → л, l → л)
    if unit_raw in ("мл", "ml"):
        unit = "л"
        # мл → литры
        try:
            v_l = float(num_raw) / 1000.0
        except Exception:
            return s
        num_raw = f"{v_l:.3f}"  # оставим 3 знака для точности, ниже обрежем
    elif unit_raw in ("л", "l", ""):
        unit = "л"
    elif unit_raw in ("кг", "kg"):
        unit = "кг"
    else:
        unit = "л"  # по умолчанию

    # приводим формат к 'x,y'
    try:
        v = float(num_raw)
    except Exception:
        return s
    # 1 или 2 знака после запятой (как было в исходнике)
    txt = f"{v:.2f}".replace(".", ",")
    # уберём лишний ноль в сотых, если ровно x,0y
    ip, fp = txt.split(",")
    if fp.endswith("0"):
        fp = fp[:-1]
    if fp == "":
        fp = "0"

    return f"{int(ip)},{fp} {unit}"




def _parse_volume_ml(vol_str: str) -> int:
    """
    Перевод строкового объёма в миллилитры для ключа каталога.
    Правила:
      • 'x л' / 'x l' → x*1000 мл
      • 'x мл' / 'x ml' → x мл
      • 'x кг' / 'x kg' → считаем 1 кг ≈ 1 л → x*1000 мл (для сопоставления)
      • просто число → трактуем как литры → x*1000 мл
    """
    if not vol_str:
        return 0
    s = str(vol_str).lower().replace("\xa0", " ").strip().replace(",", ".")
    # литры
    m = re.search(r"(\d+(?:\.\d+)?)\s*(л|l)\b", s)
    if m:
        return int(round(float(m.group(1)) * 1000))
    # миллилитры
    m = re.search(r"(\d+(?:\.\d+)?)\s*(мл|ml)\b", s)
    if m:
        return int(round(float(m.group(1))))
    # килограммы → приравниваем к литрам (для ключа)
    m = re.search(r"(\d+(?:\.\d+)?)\s*(кг|kg)\b", s)
    if m:
        return int(round(float(m.group(1)) * 1000))
    # голое число → как литры
    m = re.fullmatch(r"\d+(?:\.\d+)?", s)
    if m:
        return int(round(float(s) * 1000))
    return 0

def _cleanup_flavor(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"\s*\b\d+(?:[.,]\d+)?\s*(?:л|кг|ml|мл)\b.*$", "", s, flags=re.I)
    s = re.sub(r"\bТМ\s*«[^»]+»", "", s, flags=re.I)
    s = re.sub(r'\bTM\s*"[^"]+"', "", s, flags=re.I)
    s = _QTY_TAIL_RX.sub("", s)
    s = _QTY_DIGIT_TAIL_RX.sub("", s)
    s = re.sub(r"\s{2,}", " ", s).strip(" ,;:-—")
    return s[:1].upper() + s[1:] if s else s

# ====== распознавание входного текста ========================================
# сигнатуры
# ---- парсинг буфера: TSV / письмо / авто ------------------------------------

_NAME_QTY_RX       = re.compile(r"^(?P<name>.+?)(?:\t|\s{2,})(?P<qty>[\d\s]+)$")
_CIP_RX            = re.compile(r"^\s*(?:CIP|СIP|СИП)\s*([12])\s*$", re.I)

# объём/масса
_VOL_TOKEN_RX      = re.compile(r"\b(\d+(?:[.,]\d+)?)\s*(?:л|кг|ml|мл)\b", re.I)

# количество вида "… — 1 200 шт"
_QTY_RX            = re.compile(r"(\d[\d\s]*)\s*шт\.?\b", re.I)

# количество как «голые» цифры в хвосте: "… 1 200"
_QTY_DIGIT_TAIL_RX = re.compile(r"[-–—]?\s*(\d[\d\s]{2,})\s*$", re.I)

# удалить «… — 1 200 шт» из хвоста имени
_QTY_TAIL_RX       = re.compile(r"[-–—]?\s*\d[\d\s]*\s*шт\.?\s*$", re.I)


_HEADER_SYNONYMS: Dict[str,str] = {
    "job_id":"jobid|id|задание",
    "status":"status|статус",
    "category":"category|категория|тип",
    "name":"name|наименование|sku|продукт|товар",
    "volume":"volume|объем|объём|литраж|л",
    "quantity":"quantity|qty|кол-во|колво|количество|шт",
    "line":"line|линия|номер линии",
    "speed":"speed|скорость",
    "speed_source":"speedsource|источник скорости|источник",
    "created_at":"created|создано|дата создания",
    "updated_at":"updated|обновлено|дата обновления",
    "fact_qty":"fact|факт|выпуск",
    "progress":"progress|прогресс",
    "percent_done":"percent|процент|готовность",
    "state":"state|состояние",
    "priority":"priority|приоритет",
}

def _guess_header_mapping(headers: List[str]) -> Dict[int,str]:
    mapping: Dict[int,str] = {}
    compiled = {tgt: re.compile(rf"^(?:{syn})$", re.I) for tgt, syn in _HEADER_SYNONYMS.items()}
    for idx, h in enumerate(headers):
        h_clean = re.sub(r"\s+"," ", str(h or "")).strip().lower()
        if not h_clean: continue
        if h_clean in COL_KEYS: mapping[idx] = h_clean; continue
        for tgt, rx in compiled.items():
            if rx.match(h_clean):
                mapping[idx] = tgt; break
    return mapping

def _split_rows_by_tabs(src: str) -> List[List[str]]:
    rows = []
    for line in src.splitlines():
        if not line.strip(): continue
        rows.append(line.rstrip("\r\n").split("\t"))
    return rows

def _parse_tsv_or_csv(text: str) -> List[Dict[str,Any]]:
    rows = _split_rows_by_tabs(text)
    if not rows:
        rows = [re.split(r"\s*;\s*", ln) for ln in text.splitlines() if ln.strip()]
        if not rows: return []

    # Фильтруем служебные строки: CIP, Запуск, Вытеснение
    def _is_service_row(r: List[str]) -> bool:
        if not r or not r[0]:
            return False
        first_cell = str(r[0]).strip().lower()
        # CIP любой (CIP 1, CIP 2, CIP 3, просто CIP)
        if re.match(r'^(?:cip|сip|сип)\s*\d*$', first_cell, re.I):
            return True
        # Запуск, Вытеснение
        if first_cell in ('запуск', 'вытеснение'):
            return True
        return False
    
    original_count = len(rows)
    rows = [r for r in rows if not _is_service_row(r)]
    filtered_count = original_count - len(rows)
    
    if filtered_count > 0:
        print(f"[PARSE] Отфильтровано служебных строк: {filtered_count}, осталось: {len(rows)}")
    
    if not rows: return []

    first = rows[0]
    # Проверяем, похоже ли на заголовок
    # Если это "текст + число" (например "Сок ... 240 000"), то это НЕ заголовок
    if len(first) == 2:
        first_cell = str(first[0] or "").strip()
        second_cell = str(first[1] or "").strip()
        # Если первая ячейка - текст, вторая - число → это данные, не заголовок
        if first_cell and re.fullmatch(r"\d[\d\s]*", second_cell):
            is_header = False
        else:
            nonnum = sum(1 for x in first if not re.fullmatch(r"\d[\d\s]*", str(x or "").strip()))
            is_header = nonnum >= max(1, len(first)//2)
    else:
        nonnum = sum(1 for x in first if not re.fullmatch(r"\d[\d\s]*", str(x or "").strip()))
        is_header = nonnum >= max(1, len(first)//2)

    print(f"[PARSE] Первая строка: {first[:2] if len(first) > 2 else first}")
    print(f"[PARSE] is_header = {is_header}, всего строк для обработки: {len(rows)}")

    items: List[Dict[str,Any]] = []
    mapping: Dict[int,str] = {}
    data_rows = rows[1:] if is_header else rows
    if is_header: mapping = _guess_header_mapping([str(x) for x in first])
    
    print(f"[PARSE] Строк данных для парсинга: {len(data_rows)}")

    parsed_count = 0
    skipped_count = 0
    
    for r in data_rows:
        # вариант «Имя … Кол-во» в 1 колонке
        if len(r)==1:
            m = _NAME_QTY_RX.match(r[0].strip())
            if m:
                items.append({
                    "status":"Planned",
                    "category":"",
                    "name":m.group("name").strip(),
                    "volume":"",
                    "quantity":_coerce_numeric(m.group("qty")),
                    "line":"","speed":"","speed_source":"",
                    "created_at":"","updated_at":"",
                    "fact_qty":"","progress":"","percent_done":"","state":"",
                    "priority":"",
                    "brand":"","type":"","flavor":"",
                })
                parsed_count += 1
                continue

        if is_header and mapping:
            item: Dict[str,Any] = {k:"" for k in COL_KEYS}
            for idx, cell in enumerate(r):
                tgt = mapping.get(idx); 
                if not tgt: continue
                val = str(cell).strip()
                if tgt in _NUMERIC_COLS: val = _coerce_numeric(val)
                if tgt == "volume": val = _normalize_volume_to_str(val)
                item[tgt] = val
            item["status"] = item.get("status") or "Planned"
            items.append(item)
            parsed_count += 1
        else:
            # 2 колонки: имя + qty
            if len(r)>=2 and (r[-1] or "").strip():
                last = str(r[-1]).strip()
                name = " ".join(str(x).strip() for x in r[:-1] if str(x).strip())
                if name and re.fullmatch(r"\d[\d\s]*", last):
                    items.append({
                        "status":"Planned",
                        "category":"",
                        "name":name,
                        "volume":"",
                        "quantity":_coerce_numeric(last),
                        "line":"","speed":"","speed_source":"",
                        "created_at":"","updated_at":"",
                        "fact_qty":"","progress":"","percent_done":"","state":"",
                        "priority":"",
                        "brand":"","type":"","flavor":"",
                    })
                    parsed_count += 1
                    continue
                else:
                    skipped_count += 1
            # запасной — минимум category/name/volume/qty/line
            item: Dict[str,Any] = {k:"" for k in COL_KEYS}
            if len(r)>=1: item["category"]=str(r[0]).strip()
            if len(r)>=2: item["name"]=str(r[1]).strip()
            if len(r)>=3: item["volume"]=_normalize_volume_to_str(str(r[2]).strip())
            if len(r)>=4: item["quantity"]=_coerce_numeric(str(r[3]).strip())
            if len(r)>=5: item["line"]=str(r[4]).strip()
            item["status"]="Planned"
            items.append(item)
            parsed_count += 1
    
    print(f"[PARSE] Результат: распознано {parsed_count}, пропущено {skipped_count}, итого items: {len(items)}")
    return items

def _normalize_text_basic(text: str) -> str:
    t = text.replace("\u00A0"," ").replace("–","-").replace("—","-")
    t = re.sub(r"[ \t]+"," ", t)
    return "\n".join([ln.rstrip() for ln in t.splitlines()])

def _extract_type_flavor_brand(name_src: str, volume: str) -> tuple[str, str, str]:
    brand = ""
    mbr = re.search(r'ТМ\s*[«"]([^»"]+)[»"]', name_src, flags=re.I)
    if mbr:
        brand = mbr.group(1).strip()

    rx = re.compile(
        r'^(Сироп|Концентрат|Основа|Топпинг)\s+'
        r'(?:со вкусом и ароматом\s+)?'
        r'(?:\"([^\"]+)\"|«([^»]+)»|([^,]+?))'
        r'(?:\s+|,|$)',
        re.I
    )
    m = rx.search(name_src)
    if m:
        typ = m.group(1).capitalize() if m.group(1) else ""
        raw_flv = m.group(2) or m.group(3) or m.group(4) or ""
        flv = _cleanup_flavor(raw_flv)
    else:
        typ = ""
        flv = ""
    return typ, flv, brand

def _parse_letter_like(text: str) -> list[dict]:
    t = _normalize_text_basic(text)
    lines = [ln for ln in t.splitlines() if ln.strip()]
    out: list[dict] = []

    for ln in lines:
        # Пропускаем CIP, Запуск, Вытеснение
        ln_lower = ln.strip().lower()
        if _CIP_RX.match(ln):
            continue
        if ln_lower in ('запуск', 'вытеснение'):
            continue

        m = _NAME_QTY_RX.match(ln)
        if m:
            out.append({
                "status":"Planned","category":"",
                "name":m.group("name").strip(),
                "volume":"","quantity":_coerce_numeric(m.group("qty")),
                "line":"","speed":"","speed_source":"",
                "created_at":"","updated_at":"",
                "fact_qty":"","progress":"","percent_done":"","state":"",
                "priority":"","brand":"","type":"","flavor":"",
            })
            continue

        name = _QTY_TAIL_RX.sub("", ln).strip()
        mqt = _QTY_DIGIT_TAIL_RX.search(name)
        qty = ""
        if mqt:
            qty = _coerce_numeric(mqt.group(1))
            name = _QTY_DIGIT_TAIL_RX.sub("", name).strip()

        vol = ""
        mv = _VOL_TOKEN_RX.search(name)
        if mv:
            vol = _normalize_volume_to_str(mv.group(0))

        out.append({
            "status":"Planned","category":"",
            "name":name,"volume":vol,"quantity":qty,
            "line":"","speed":"","speed_source":"",
            "created_at":"","updated_at":"",
            "fact_qty":"","progress":"","percent_done":"","state":"",
            "priority":"","brand":"","type":"","flavor":"",
        })
    return out

def _row_score(r: Dict[str,Any]) -> int:
    nm = (r.get("name") or "").strip()
    hints = sum(bool(r.get(k)) for k in ("volume","quantity","line"))
    looks = bool(nm) and (re.search(r"\bТМ\b",nm) or re.search(r"\b(сироп|концентрат|основа|топпинг)\b", nm, re.I))
    return 1 if looks or hints>=2 else 0

def _score_rows(rows: List[Dict[str,Any]]) -> int:
    return sum(_row_score(r) for r in rows)

def parse_clipboard_text(src: str) -> Tuple[List[Dict[str,Any]], str]:
    s = src.strip()
    if not s: return [], "empty"
    rows_tsv = _parse_tsv_or_csv(s)
    rows_let = _parse_letter_like(s)
    sc_tsv, sc_let = _score_rows(rows_tsv), _score_rows(rows_let)

    if sc_tsv >= sc_let and rows_tsv: return rows_tsv, "Excel-TSV"
    if rows_let: return rows_let, "Письмо"
    if rows_tsv: return rows_tsv, "CSV/;"
    return [], "unknown"

# ====== Каталог: загрузка/поиск/добавление ===================================
_catalog_by_name: dict[str, dict] = {}
_catalog_by_key: dict[str, dict] = {}

def _norm_name_match(s: str) -> str:
    """Нормализация имени для сопоставления: убираем объем, кавычки, пробелы"""
    s = str(s or "").replace("\u00A0"," ")
    s = s.replace("«",'"').replace("»",'"')
    # Убираем объем/массу из имени: "0,25 л", "1,0 кг", "250 мл" и т.д.
    s = re.sub(r'\b\d+[.,]?\d*\s*(?:л|l|кг|kg|мл|ml)\b', '', s, flags=re.I)
    s = re.sub(r"\s+"," ", s).strip()
    return s.lower()

def _load_catalog_maps() -> None:
    global _catalog_by_name, _catalog_by_key
    if _catalog_by_name or _catalog_by_key: return
    path = _catalog_path()
    try:
        data = json.load(open(path,"r",encoding="utf-8"))
    except Exception:
        _catalog_by_name=_catalog_by_key={}
        return
    if not isinstance(data, list): return
    _catalog_by_name, _catalog_by_key = {}, {}
    for row in data:
        if not isinstance(row, dict): continue
        nm = _norm_name_match(row.get("name",""))
        if nm: _catalog_by_name[nm] = row
        # ключ по продукт-парсеру
        if _pparse:
            try:
                pp = _pparse(row.get("name",""), row.get("container",""))
                typ = (pp.get("type") or "").strip().lower()
                flv = (pp.get("flavor") or "").strip().lower()
                brd = (pp.get("brand") or "").strip().lower()
                vml = _parse_volume_ml(row.get("container",""))
                key = f"{typ}|{flv}|{brd}|{vml}"
                if typ or flv: _catalog_by_key[key] = {
                    "speed": row.get("speed", None),
                    "speed_source": "Каталог" if row.get("speed") not in (None,"") else "",
                    "line_default": row.get("line",""),
                }
            except Exception:
                pass

def _product_key(name: str, volume: str) -> str:
    if not _pparse: return ""
    try:
        pp = _pparse(name, volume)
        typ = (pp.get("type") or "").strip().lower()
        flv = (pp.get("flavor") or "").strip().lower()
        brd = (pp.get("brand") or "").strip().lower()
        vml = _parse_volume_ml(volume)
        return f"{typ}|{flv}|{brd}|{vml}"
    except Exception:
        return ""

def _catalog_match_status(name: str, volume: str) -> str:
    _load_catalog_maps()
    if _norm_name_match(name) in _catalog_by_name: return "exact"
    if _pparse and _product_key(name, volume) in _catalog_by_key: return "partial"
    return "none"

_SOURCE_STRENGTH = {"матрица":4,"норматив":3,"замер":3,"история":2,"оценка":1,"каталог":2}
def _strength(src: str) -> int:
    return _SOURCE_STRENGTH.get(str(src or "").strip().lower(), 0)

def _enrich_from_catalog(row: dict, preserve_line_if_set: bool, overwrite_speed_if_stronger: bool) -> dict:
    _load_catalog_maps()
    out = dict(row)

    rec = _catalog_by_name.get(_norm_name_match(out.get("name","")))
    if not rec and _pparse:
        rec = _catalog_by_key.get(_product_key(out.get("name",""), out.get("volume","")))

    if rec:
        line_def = rec.get("line") if "line" in rec else rec.get("line_default")
        if line_def and not (out.get("line") and preserve_line_if_set):
            out["line"] = str(line_def)

        spd = rec.get("speed"); src = rec.get("speed_source") or ("Каталог" if spd not in (None,"") else "")
        if out.get("speed"):
            if overwrite_speed_if_stronger and _strength(src) > _strength(out.get("speed_source","")):
                out["speed"] = str(spd) if spd not in (None,"") else out.get("speed","")
                out["speed_source"] = src or out.get("speed_source","")
        else:
            if spd not in (None,""):
                out["speed"] = str(spd); out["speed_source"] = src

    return out

def _append_to_catalog_by_name(entries: list[dict]) -> int:
    path = _catalog_path()
    data: list = []
    if os.path.isfile(path):
        try:
            data = json.load(open(path,"r",encoding="utf-8"))
            if not isinstance(data,list): data=[]
        except Exception:
            data=[]
    def nkey(n: str) -> str: return _norm_name_match(n)
    idx = {nkey(r.get("name","")): i for i,r in enumerate(data) if isinstance(r,dict) and r.get("name")}
    added = 0
    for e in entries:
        nm = e.get("name",""); 
        if not nm: continue
        key = nkey(nm)
        if key in idx:
            i = idx[key]; rec = data[i]
            for fld in ("line","container","speed","limit","action"):
                if rec.get(fld) in (None,"",0) and e.get(fld) not in (None,""):
                    rec[fld] = e[fld]
            data[i] = rec
        else:
            data.append({
                "name": nm,
                "line": e.get("line",""),
                "container": e.get("container",""),
                "speed": e.get("speed", None),
                "limit": e.get("limit", None),
                "action": e.get("action",""),
            })
            idx[key] = len(data)-1; added += 1
    with open(path,"w",encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    global _catalog_by_name, _catalog_by_key
    _catalog_by_name=_catalog_by_key={}
    _load_catalog_maps()
    return added

# ====== ПУБЛИЧНАЯ ТОЧКА ВХОДА ================================================
def show_planning_tab(nb: ttk.Notebook):
    # убрать старую
    try:
        for tid in list(nb.tabs()):
            if nb.tab(tid, "text") == "Планирование":
                nb.forget(tid)
    except Exception: pass

    tab_planning = ttk.Frame(nb); nb.add(tab_planning, text="Планирование")
    sub = ttk.Notebook(tab_planning); sub.pack(fill="both", expand=True)
    tab_plan  = ttk.Frame(sub); sub.add(tab_plan, text="План")
    tab_sched = ttk.Frame(sub); sub.add(tab_sched, text="Расписание")
    tab_fact  = ttk.Frame(sub); sub.add(tab_fact, text="Факт/План")
    tab_import= ttk.Frame(sub); sub.add(tab_import, text="Импорт")

    # ---------- ПЛАН ----------
    # === ВЕРХНЯЯ ПАНЕЛЬ УПРАВЛЕНИЯ ===
    control_frame = ttk.Frame(tab_plan)
    control_frame.pack(fill="x", padx=8, pady=(8, 4))
    
    # Левая группа - управление записями
    left_group = ttk.LabelFrame(control_frame, text="Управление записями", padding=8)
    left_group.pack(side="left", fill="x", expand=True, padx=(0, 8))
    
    btn_add = ttk.Button(left_group, text="➕ Добавить")
    btn_add.pack(side="left", padx=(0, 6))
    
    btn_dup = ttk.Button(left_group, text="📋 Дублировать")
    btn_dup.pack(side="left", padx=(0, 6))
    
    btn_del = ttk.Button(left_group, text="🗑️ Удалить")
    btn_del.pack(side="left", padx=(0, 12))
    
    btn_enrich = ttk.Button(left_group, text="✨ Обогатить из каталога")
    btn_enrich.pack(side="left", padx=(0, 6))
    
    btn_lock_priorities = ttk.Button(left_group, text="🔒 Заблокировать приоритеты")
    btn_lock_priorities.pack(side="left", padx=(0, 6))

    btn_change_status = ttk.Button(left_group, text="📝 Изменить статус")
    btn_change_status.pack(side="left", padx=(0, 6))
    
    # Правая группа - файловые операции
    right_group = ttk.LabelFrame(control_frame, text="Файл", padding=8)
    right_group.pack(side="right", fill="x")
    
    btn_load = ttk.Button(right_group, text="📂 Загрузить")
    btn_load.pack(side="left", padx=(0, 6))
    
    btn_save = ttk.Button(right_group, text="💾 Сохранить")
    btn_save.pack(side="left", padx=(0, 6))
    
    # === ОПЦИИ ОБОГАЩЕНИЯ ===
    options_frame = ttk.Frame(tab_plan)
    options_frame.pack(fill="x", padx=8, pady=(0, 4))
    
    var_preserve_line = tk.BooleanVar(value=True)
    var_overwrite_speed = tk.BooleanVar(value=False)
    
    ttk.Checkbutton(options_frame, text="Сохранять линию при обогащении", 
                   variable=var_preserve_line).pack(side="left", padx=(0, 12))
    ttk.Checkbutton(options_frame, text="Перезаписывать скорость из каталога", 
                   variable=var_overwrite_speed).pack(side="left")
    
    # === ОСНОВНАЯ ТАБЛИЦА ===
    table_frame = ttk.Frame(tab_plan)
    table_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    
    tree = ttk.Treeview(table_frame, columns=COL_KEYS, show="tree headings", 
                       selectmode="extended", height=20)
    
    # Настройка скроллбаров
    scY = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    scX = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=scY.set, xscrollcommand=scX.set)
    
    # Размещение с использованием grid
    tree.grid(row=0, column=0, sticky="nsew")
    scY.grid(row=0, column=1, sticky="ns")
    scX.grid(row=1, column=0, sticky="ew")
    
    table_frame.grid_rowconfigure(0, weight=1)
    table_frame.grid_columnconfigure(0, weight=1)
    
    # Настройка колонок с улучшенными заголовками
    _config_tree(tree, COL_KEYS, COL_HEADERS, COL_WIDTHS, _NUMERIC_COLS)
    
    # Настройка стилей для строк
    tree.tag_configure("row_odd", background="#f8f9fa")
    tree.tag_configure("row_even", background="#ffffff")
    tree.tag_configure("completed", background="#d4edda", foreground="#155724")
    tree.tag_configure("in_progress", background="#fff3cd", foreground="#856404")
    tree.tag_configure("planned", background="#ffffff")
    tree.tag_configure("postponed", background="#f8d7da", foreground="#721c24")
    
    # === СТАТУСНАЯ СТРОКА ===
    status_frame = ttk.Frame(tab_plan)
    status_frame.pack(fill="x", padx=8, pady=(0, 8))
    
    info_lbl = ttk.Label(status_frame, text="Готов к работе", foreground="#666666")
    info_lbl.pack(side="left")
    
    # Счетчик записей
    count_lbl = ttk.Label(status_frame, text="", foreground="#007bff")
    count_lbl.pack(side="right")

    def _update_group_count(parent):
        """Обновление счетчика в заголовке группы линии"""
        children_count = len(tree.get_children(parent))
        line_name = tree.item(parent, "text")
        # Убираем старый счетчик если есть
        line_name = line_name.split(" (")[0]
        tree.item(parent, text=f"{line_name} ({children_count})")
    
    def _update_count():
        """Обновление счетчика записей"""
        total = 0
        for parent in tree.get_children(""):
            total += len(tree.get_children(parent))
        if total > 0:
            count_lbl.config(text=f"Всего записей: {total}")
        else:
            count_lbl.config(text="")
    
    def _insert(values: Dict[str,Any]) -> str:
        """Вставка записи с группировкой по линиям"""
        vals = [values.get(k,"") for k in COL_KEYS]
        
        # Определяем тег по статусу
        status = values.get("status", "").lower()
        if "complete" in status or "завершен" in status:
            tag = "completed"
        elif "progress" in status or "выполн" in status or "in progress" in status:
            tag = "in_progress"
        elif "postponed" in status or "отложен" in status:
            tag = "postponed"
        else:
            tag = "planned"
        
        # Определяем линию для группировки
        line = values.get("line", "").strip() or "Без линии"
        
        # Ищем существующую группу линии
        parent = None
        for item in tree.get_children(""):
            if tree.item(item, "text").startswith(f"📍 {line}"):
                parent = item
                break
        
        # Если группы нет, создаем
        if parent is None:
            parent = tree.insert("", "end", text=f"📍 {line}", values=("",) * len(COL_KEYS))
            tree.item(parent, open=True)
            _update_group_count(parent)
        else:
            _update_group_count(parent)
        
        # Без какой-либо сортировки — вставляем строго в конец, сохраняя исходный порядок
        iid = tree.insert(parent, "end", values=tuple(vals), tags=(tag,))
        
        # Автоподгонка ширины столбцов после добавления
        _autofit_columns(tree)
        
        _update_count()
        return iid

    def _load_json(path: str=_PLAN_JSON):
        """Улучшенная загрузка с обновлением статуса"""
        try:
            if not os.path.isfile(path):
                info_lbl.config(text="📄 Файл не найден — начните заполнять")
                count_lbl.config(text="")
                return
            
            with open(path, "r", encoding="utf-8") as f:
                rows = json.load(f)
            
            tree.delete(*tree.get_children(""))
            for r in rows:
                _insert(r)
            
            # Автоподгонка ширины столбцов
            _autofit_columns(tree)
            
            info_lbl.config(text=f"✅ Загружено из {os.path.basename(path)}")
            _update_count()
            
        except Exception as e:
            messagebox.showerror("Загрузка плана", f"Не удалось загрузить:\n{e}")
            info_lbl.config(text="❌ Ошибка загрузки")

    def _save_json(path: str=_PLAN_JSON):
        """Улучшенное сохранение с обновлением статуса и группировкой"""
        try:
            rows: List[Dict[str,Any]] = []
            for iid in tree.get_children(""):
                if tree.item(iid, "text").startswith("📍"):
                    # Группа линий - собираем дочерние элементы
                    for child in tree.get_children(iid):
                        vals = tree.item(child, "values")
                        rows.append({k: (vals[i] if i < len(vals) else "") for i, k in enumerate(COL_KEYS)})
                else:
                    # Прямая запись
                    vals = tree.item(iid, "values")
                    rows.append({k: (vals[i] if i < len(vals) else "") for i, k in enumerate(COL_KEYS)})
            
            # Сохраняем с временным файлом для безопасности
            temp_path = path + ".tmp"
            with open(temp_path, "w", encoding="utf-8") as f:
                json.dump(rows, f, ensure_ascii=False, indent=2)
            
            # Заменяем оригинальный файл
            import shutil
            shutil.move(temp_path, path)
            
            info_lbl.config(text=f"✅ Сохранено в {os.path.basename(path)}")
            _update_count()
            
        except Exception as e:
            messagebox.showerror("Сохранение плана", f"Не удалось сохранить:\n{e}")
            info_lbl.config(text="❌ Ошибка сохранения")

    if os.path.isfile(_PLAN_JSON): _load_json()
    else:
        _insert({
            "job_id":"J-250915-L01-001","status":"Planned","category":"Сироп",
            "name":"Сироп со вкусом и ароматом \"Ваниль\" ТМ «Пример»",
            "volume":"1,0 л","quantity":"1500","line":"Линия 1",
            "speed":"1100","speed_source":"Матрица",
            "created_at":"","updated_at":"",
            "fact_qty":"0","progress":"0 / 1500","percent_done":"0,0%","state":"Не начато",
            "priority":"3","flavor":"Ваниль","brand":"Пример","type":"Сироп",
        })

    # редактирование ячеек
    _ed_entry: Optional[tk.Entry]=None; _ed_item: Optional[str]=None; _ed_col: Optional[str]=None
    def _bbox(item, col):
        try: 
            b = tree.bbox(item,col); 
            return b if b else None
        except Exception: return None
    def _start_edit(_e=None):
        nonlocal _ed_entry,_ed_item,_ed_col
        if tree.identify("region", _e.x, _e.y)!="cell": return
        col = tree.identify_column(_e.x); row = tree.identify_row(_e.y)
        if not row or not col: return
        bx=_bbox(row,col); 
        if not bx: return
        x,y,w,h=bx; col_idx=int(col[1:])-1; col_name=tree["columns"][col_idx]
        cur=tree.set(row,col_name)
        _ed_item,_ed_col=row,col
        _ed_entry=tk.Entry(tree); _ed_entry.insert(0,cur); _ed_entry.select_range(0,"end")
        _ed_entry.focus_set(); _ed_entry.place(x=x,y=y,width=w,height=h)
        def _commit(e=None):
            val=_ed_entry.get()
            if col_name in _NUMERIC_COLS: val=_coerce_numeric(val)
            tree.set(_ed_item,col_name,val)
            
            # Автосортировка ОТКЛЮЧЕНА - используйте drag & drop или кнопку "Сортировать"
            # if col_name == "priority":
            #     _sort_item_by_priority(_ed_item)
            
            _cancel()
        def _cancel(e=None):
            nonlocal _ed_entry,_ed_item,_ed_col
            if _ed_entry: _ed_entry.destroy()
            _ed_entry=_ed_item=_ed_col=None
        _ed_entry.bind("<Return>",_commit); _ed_entry.bind("<Escape>",_cancel); _ed_entry.bind("<FocusOut>",_commit)
    tree.bind("<Double-1>", _start_edit); tree.bind("<Return>", _start_edit)

    def _add_row():
        """Добавление новой записи"""
        # Создаем пустую запись с дефолтными значениями
        new_values = {k: "" for k in COL_KEYS}
        new_values["status"] = "Planned"
        new_values["priority"] = "5"
        
        iid = _insert(new_values)
        tree.see(iid)
        tree.selection_set(iid)
        info_lbl.config(text="➕ Добавлена новая запись")
        _update_count()
    
    def _dup_rows():
        """Дублирование выбранных записей"""
        sels = tree.selection()
        if not sels:
            messagebox.showinfo("Дублирование", "Выберите записи для дублирования")
            return
        
        new = []
        for iid in sels:
            # Пропускаем группы линий
            if tree.item(iid, "text").startswith("📍"):
                continue
            vals = tree.item(iid, "values")
            # Создаем словарь из значений
            row_dict = {k: (vals[i] if i < len(vals) else "") for i, k in enumerate(COL_KEYS)}
            # Очищаем job_id для новой записи
            row_dict["job_id"] = ""
            new_iid = _insert(row_dict)
            new.append(new_iid)
        
        if new:
            tree.see(new[-1])
            tree.selection_set(new[-1])
            info_lbl.config(text=f"📋 Дублировано записей: {len(new)}")
            _update_count()
    
    def _del_rows():
        """Удаление выбранных записей"""
        sels = tree.selection()
        if not sels:
            messagebox.showinfo("Удаление", "Выберите записи для удаления")
            return
        
        # Удаляем только записи, не группы
        to_delete = [iid for iid in sels if not tree.item(iid, "text").startswith("📍")]
        
        count = len(to_delete)
        if count == 0:
            messagebox.showinfo("Удаление", "Выберите записи, а не группы линий")
            return
        
        if count > 1:
            if not messagebox.askyesno("Удаление", f"Удалить {count} записей?"):
                return
        
        for iid in to_delete:
            tree.delete(iid)
        
        # Удаляем пустые группы и обновляем счетчики
        for item in list(tree.get_children("")):
            if tree.item(item, "text").startswith("📍"):
                if not tree.get_children(item):
                    tree.delete(item)
                else:
                    _update_group_count(item)
        
        info_lbl.config(text=f"🗑️ Удалено записей: {count}")
        _update_count()

    def _sort_item_by_priority(item_id):
        """Сортировка конкретного элемента по приоритету"""
        try:
            parent = tree.parent(item_id)
            if not parent:
                return  # Элемент не в группе
            
            # Получаем новый приоритет
            vals = tree.item(item_id, "values")
            new_priority = int(vals[0]) if vals and vals[0] else 999
            
            # Удаляем элемент
            tree.delete(item_id)
            
            # Находим правильную позицию для вставки
            insert_pos = len(tree.get_children(parent))
            for idx, child in enumerate(tree.get_children(parent)):
                child_vals = tree.item(child, "values")
                child_priority = int(child_vals[0]) if child_vals and child_vals[0] else 999
                if new_priority < child_priority:
                    insert_pos = idx
                    break
            
            # Вставляем обратно в правильную позицию с сохранением тега статуса
            status = vals[COL_KEYS.index("status")].lower() if len(vals) > COL_KEYS.index("status") else ""
            if "complete" in status or "завершен" in status:
                tag = "completed"
            elif "progress" in status or "выполн" in status or "in progress" in status:
                tag = "in_progress"
            elif "postponed" in status or "отложен" in status:
                tag = "postponed"
            else:
                tag = "planned"
            new_item = tree.insert(parent, insert_pos, values=vals, tags=(tag,))
            tree.see(new_item)
            tree.selection_set(new_item)
            
        except Exception as e:
            print(f"Ошибка сортировки: {e}")

    def _sort_all_by_priority():
        """Полная сортировка всех записей по приоритету"""
        try:
            # Собираем все записи
            all_records = []
            for parent in tree.get_children(""):
                if tree.item(parent, "text").startswith("📍"):
                    # Группа линий - собираем дочерние элементы
                    for child in tree.get_children(parent):
                        vals = tree.item(child, "values")
                        record = {k: (vals[i] if i < len(vals) else "") for i, k in enumerate(COL_KEYS)}
                        record["_parent"] = parent
                        record["_item"] = child
                        all_records.append(record)
                else:
                    # Прямая запись
                    vals = tree.item(parent, "values")
                    record = {k: (vals[i] if i < len(vals) else "") for i, k in enumerate(COL_KEYS)}
                    record["_parent"] = None
                    record["_item"] = parent
                    all_records.append(record)
            
            # Сортируем по приоритету
            sorted_records = sorted(all_records, key=lambda r: int(r.get("priority", 999) or 999))
            
            # Очищаем дерево
            tree.delete(*tree.get_children(""))
            
            # Вставляем отсортированные записи
            for record in sorted_records:
                del record["_parent"]
                del record["_item"]
                _insert(record)
            
            info_lbl.config(text=f"🎯 Отсортировано {len(sorted_records)} записей по приоритету")
            _update_count()
            
        except Exception as e:
            messagebox.showerror("Ошибка сортировки", f"Не удалось отсортировать:\n{e}")

    def _open_lock_priorities_window():
        """Открытие окна блокировки приоритетов - простой подход"""
        import tkinter as tk
        from tkinter import ttk, messagebox
        
        # Читаем данные из файла напрямую
        try:
            with open("jobs_plan.json", "r", encoding="utf-8") as f:
                jobs_data = json.load(f)
        except:
            messagebox.showerror("Ошибка", "Не удалось загрузить jobs_plan.json")
            return
        
        # Находим все приоритеты
        priorities = set()
        for job in jobs_data:
            priority = job.get("priority", "")
            if priority and priority.isdigit():
                priorities.add(int(priority))
        
        if not priorities:
            messagebox.showinfo("Информация", "В файле нет приоритетов")
            return
        
        priorities = sorted(priorities)
        
        # Загружаем текущие настройки
        locked_priorities = set()
        try:
            with open("locked_priorities.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                locked_priorities = set(data.get("locked", []))
        except:
            pass
        
        # Создаем окно
        window = tk.Toplevel()
        window.title("Блокировка приоритетов")
        window.geometry("400x300")
        window.transient()
        window.grab_set()
        
        # Заголовок
        ttk.Label(window, text="Выберите приоритеты для блокировки", 
                 font=("Arial", 12, "bold")).pack(pady=10)
        
        # Информация
        info_text = f"Найдено {len(priorities)} групп приоритетов в {len(jobs_data)} заданиях"
        ttk.Label(window, text=info_text, foreground="#666").pack(pady=5)
        
        # Фрейм для чекбоксов
        frame = ttk.Frame(window)
        frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Создаем чекбоксы
        vars_dict = {}
        for priority in priorities:
            count = len([j for j in jobs_data if j.get("priority") == str(priority)])
            
            var = tk.BooleanVar(value=priority in locked_priorities)
            vars_dict[priority] = var
            
            cb = ttk.Checkbutton(frame, 
                               text=f"Приоритет {priority} ({count} заданий)",
                               variable=var)
            cb.pack(anchor="w", pady=2)
        
        # Кнопки
        btn_frame = ttk.Frame(window)
        btn_frame.pack(fill="x", padx=20, pady=10)
        
        def save_and_close():
            # Собираем заблокированные приоритеты
            locked = [p for p, var in vars_dict.items() if var.get()]
            
            # Сохраняем
            try:
                with open("locked_priorities.json", "w", encoding="utf-8") as f:
                    json.dump({"locked": locked}, f, ensure_ascii=False, indent=2)
                
                print(f"Заблокировано {len(locked)} приоритетов: {locked}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}")
                return
            
            window.destroy()
        
        ttk.Button(btn_frame, text="💾 Сохранить", command=save_and_close).pack(side="left", padx=(0, 10))
        ttk.Button(btn_frame, text="❌ Отмена", command=window.destroy).pack(side="left")

    def _change_status_selected():
        """Изменение статуса выбранных записей"""
        sels = tree.selection()
        if not sels:
            messagebox.showinfo("Изменение статуса", "Выберите записи для изменения статуса")
            return

        # Пропускаем группы линий
        valid_sels = [iid for iid in sels if not tree.item(iid, "text").startswith("📍")]
        if not valid_sels:
            messagebox.showinfo("Изменение статуса", "Выберите записи, а не группы линий")
            return

        # Создаем диалог выбора статуса
        status_window = tk.Toplevel(tab_plan)
        status_window.title("Изменить статус")
        status_window.geometry("300x200")
        status_window.transient(tab_plan)
        status_window.grab_set()

        ttk.Label(status_window, text="Выберите новый статус:", font=("", 11, "bold")).pack(pady=10)

        # Переменная для выбора статуса
        status_var = tk.StringVar(value="Planned")

        # Радиокнопки для разных статусов
        ttk.Radiobutton(status_window, text="Запланировано", variable=status_var, value="Planned").pack(anchor="w", padx=20)
        ttk.Radiobutton(status_window, text="Отложено", variable=status_var, value="Postponed").pack(anchor="w", padx=20)
        ttk.Radiobutton(status_window, text="В работе", variable=status_var, value="In Progress").pack(anchor="w", padx=20)
        ttk.Radiobutton(status_window, text="Завершено", variable=status_var, value="Completed").pack(anchor="w", padx=20)

        def apply_status():
            new_status = status_var.get()
            updated = 0

            for iid in valid_sels:
                vals = list(tree.item(iid, "values"))
                if len(vals) > COL_KEYS.index("status"):
                    vals[COL_KEYS.index("status")] = new_status

                    # Определяем тег по статусу для визуального отображения
                    if new_status == "Completed":
                        tag = "completed"
                    elif new_status == "In Progress":
                        tag = "in_progress"
                    elif new_status == "Postponed":
                        tag = "postponed"
                    else:
                        tag = "planned"

                    tree.item(iid, values=tuple(vals), tags=(tag,))
                    updated += 1

            if updated > 0:
                info_lbl.config(text=f"📝 Изменен статус у {updated} записей")
                _save_json()  # Автосохранение
                _update_count()

            status_window.destroy()

        # Кнопки управления
        btn_frame = ttk.Frame(status_window)
        btn_frame.pack(fill="x", pady=20)

        ttk.Button(btn_frame, text="Применить", command=apply_status).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Отмена", command=status_window.destroy).pack(side="right", padx=10)

        # Центрируем окно
        status_window.update_idletasks()
        x = tab_plan.winfo_rootx() + (tab_plan.winfo_width() // 2) - (status_window.winfo_width() // 2)
        y = tab_plan.winfo_rooty() + (tab_plan.winfo_height() // 2) - (status_window.winfo_height() // 2)
        status_window.geometry(f"+{x}+{y}")

    def _enrich_plan():
        """Обогащение данных из каталога с улучшенным интерфейсом"""
        sels = tree.selection()
        if not sels:
            # Если ничего не выбрано, обогащаем все записи
            if not messagebox.askyesno("Обогащение", 
                                      "Ничего не выбрано. Обогатить все записи из каталога?"):
                return
            # Собираем все записи из всех групп
            all_items = []
            for parent in tree.get_children(""):
                all_items.extend(tree.get_children(parent))
            sels = all_items
        
        updated = 0
        skipped = 0
        preserve_line = bool(var_preserve_line.get())
        overwrite_speed = bool(var_overwrite_speed.get())
        
        for iid in sels:
            # Пропускаем группы линий
            if tree.item(iid, "text").startswith("📍"):
                continue
            vals = list(tree.item(iid, "values"))
            base = {k: (vals[i] if i < len(COL_KEYS) else "") for i, k in enumerate(COL_KEYS)}
            
            try:
                enr = _enrich_from_catalog(base, 
                                          preserve_line_if_set=preserve_line,
                                          overwrite_speed_if_stronger=overwrite_speed)
                if enr != base:
                    updated += 1
                    # Обновляем с сохранением стиля по статусу
                    status = enr.get("status", "").lower()
                    if "complete" in status or "завершен" in status:
                        tag = "completed"
                    elif "progress" in status or "выполн" in status or "in progress" in status:
                        tag = "in_progress"
                    elif "postponed" in status or "отложен" in status:
                        tag = "postponed"
                    else:
                        tag = "planned"
                    
                    tree.item(iid, values=tuple(enr.get(k, "") for k in COL_KEYS), tags=(tag,))
                else:
                    skipped += 1
            except Exception as e:
                print(f"Ошибка обогащения записи: {e}")
                skipped += 1
        
        # Автоподгонка после обогащения
        if updated > 0:
            _autofit_columns(tree)
        
        if updated > 0:
            info_lbl.config(text=f"✨ Обогащено из каталога: {updated} записей (пропущено: {skipped})")
        else:
            info_lbl.config(text=f"ℹ️ Нет изменений. Проверьте наличие данных в каталоге")

    # === КОНТЕКСТНОЕ МЕНЮ ===
    def _show_context_menu(event):
        """Показать контекстное меню при правом клике"""
        item = tree.identify_row(event.y)
        if item:
            tree.selection_set(item)
        
        menu = tk.Menu(tab_plan, tearoff=0)
        menu.add_command(label="➕ Добавить", command=_add_row, accelerator="Ctrl+N")
        menu.add_command(label="📋 Дублировать", command=_dup_rows, accelerator="Ctrl+D")
        menu.add_command(label="🗑️ Удалить", command=_del_rows, accelerator="Delete")
        menu.add_separator()
        menu.add_command(label="✨ Обогатить", command=_enrich_plan, accelerator="Ctrl+E")
        menu.add_separator()
        menu.add_command(label="📝 Изменить статус", command=_change_status_selected, accelerator="Ctrl+T")
        menu.add_separator()
        menu.add_command(label="💾 Сохранить", command=lambda: _save_json(), accelerator="Ctrl+S")
        
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    
    # === ГОРЯЧИЕ КЛАВИШИ ===
    def _setup_hotkeys():
        """Настройка горячих клавиш"""
        tab_plan.bind_all("<Control-n>", lambda e: _add_row())
        tab_plan.bind_all("<Control-d>", lambda e: _dup_rows())
        tab_plan.bind_all("<Delete>", lambda e: _del_rows())
        tab_plan.bind_all("<Control-s>", lambda e: _save_json())
        tab_plan.bind_all("<Control-e>", lambda e: _enrich_plan())
        tab_plan.bind_all("<Control-t>", lambda e: _change_status_selected())
        tab_plan.bind_all("<Control-l>", lambda e: _load_json())
    
    # Привязка событий
    tree.bind("<Button-3>", _show_context_menu)
    _setup_hotkeys()
    
    # Настройка команд кнопок
    btn_add.configure(command=_add_row)
    btn_dup.configure(command=_dup_rows)
    btn_del.configure(command=_del_rows)
    btn_save.configure(command=lambda: _save_json())
    btn_load.configure(command=lambda: _load_json())
    btn_enrich.configure(command=_enrich_plan)
    btn_change_status.configure(command=_change_status_selected)
    btn_lock_priorities.configure(command=_open_lock_priorities_window)
    
    # Включаем Drag & Drop для перетаскивания строк
    def _on_drag_reorder():
        """Callback после перетаскивания - автосохранение"""
        _save_json()
        info_lbl.config(text="✅ Порядок изменен и сохранен")
    
    _enable_drag_and_drop(tree, on_reorder_callback=_on_drag_reorder)
    info_lbl.config(text="🖱 Перетаскивайте строки мышью для изменения порядка")

    # ---------- РАСПИСАНИЕ ----------
    # Создаем Notebook внутри tab_sched для подвкладок "Расписание" и "Импорт JSON"
    sched_notebook = ttk.Notebook(tab_sched)
    sched_notebook.pack(fill="both", expand=True)
    
    tab_schedule_main = ttk.Frame(sched_notebook)
    tab_schedule_import = ttk.Frame(sched_notebook)
    
    sched_notebook.add(tab_schedule_main, text="Расписание")
    sched_notebook.add(tab_schedule_import, text="Импорт JSON")
    
    # Вкладка "Расписание"
    try:
        from schedule_tab import ScheduleTab
        ScheduleTab(tab_schedule_main)
    except Exception as e:
        import traceback
        traceback.print_exc()
        ttk.Label(tab_schedule_main, text=f"Ошибка при инициализации расписания: {e}", foreground="#a00")\
           .pack(anchor="w", padx=8, pady=8)
    
    # Вкладка "Импорт JSON" внутри Расписания
    try:
        from json_import_tab import JsonImportTab
        # Передаем реальный nb (главный Notebook) для доступа к вкладке Планирование
        # и уже созданную вкладку tab_schedule_import как parent_frame
        JsonImportTab(nb, parent_frame=tab_schedule_import)
    except Exception as e:
        import traceback
        traceback.print_exc()
        ttk.Label(tab_schedule_import, text=f"Ошибка при инициализации импорта: {e}", foreground="#a00")\
           .pack(anchor="w", padx=8, pady=8)
    
    # ---------- ФАКТ/ПЛАН ----------
    try:
        from fact_comparison_tab import FactComparisonTab
        FactComparisonTab(tab_fact, parent_notebook=nb)
    except Exception as e:
        import traceback
        traceback.print_exc()
        ttk.Label(tab_fact, text=f"Ошибка при инициализации сравнения: {e}", foreground="#a00")\
           .pack(anchor="w", padx=8, pady=8)

    # ---------- ИМПОРТ ----------
    top_imp = ttk.Frame(tab_import); top_imp.pack(fill="x", padx=8, pady=(8,4))
    btn_clip  = ttk.Button(top_imp, text="📋 Вставить из буфера")
    btn_clear = ttk.Button(top_imp, text="🗑 Очистить")
    btn_parse = ttk.Button(top_imp, text="🧩 Распознать →")
    btn_enrich= ttk.Button(top_imp, text="↔ Сопоставить с каталогом")
    btn_show_miss = ttk.Button(top_imp, text="🔍 Несовпадения")
    btn_addcat= ttk.Button(top_imp, text="＋ Добавить в каталог (выбранные)")
    # Массовая установка линии для распознанных продуктов
    btn_set_line_sel = ttk.Button(top_imp, text="Линия → выбранным")
    btn_set_line_all = ttk.Button(top_imp, text="Линия → всем")
    # Массовая установка скорости
    btn_set_speed_sel = ttk.Button(top_imp, text="Скорость → выбранным")
    btn_set_speed_all = ttk.Button(top_imp, text="Скорость → всем")
    btn_apply = ttk.Button(top_imp, text="⬇ Импортировать в План (Ctrl+Enter)")
    lbl_info  = ttk.Label(top_imp, text="", foreground="#666")
    for b,p in [
        (btn_clip,0),(btn_clear,6),(btn_parse,12),(btn_enrich,6),
        (btn_show_miss,6),(btn_addcat,6),(btn_set_line_sel,12),(btn_set_line_all,6),
        (btn_set_speed_sel,12),(btn_set_speed_all,6),
        (btn_apply,12)
    ]:
        b.pack(side="left", padx=p)
    lbl_info.pack(side="left", padx=12)

    opts = ttk.Frame(tab_import); opts.pack(fill="x", padx=8, pady=(0,4))
    var_preserve_line = tk.BooleanVar(value=True)
    var_overwrite_speed = tk.BooleanVar(value=False)
    ttk.Checkbutton(opts, text="Не менять line, если уже задана", variable=var_preserve_line).pack(side="left")
    ttk.Checkbutton(opts, text="Перезаписывать speed, если источник сильнее", variable=var_overwrite_speed)\
        .pack(side="left", padx=12)

    txt, tree_imp, _ = _create_import_panes(tab_import, top_imp)

    # Колонки для импорта (только нужные)
    IMP_COLS = ("name", "volume", "quantity", "type", "flavor", "brand", "line", "speed", "cat_match")
    IMP_HEADERS = ("Наименование", "Объём", "Кол-во", "Тип", "Вкус", "Бренд", "Линия", "Скорость", "Каталог")
    IMP_WIDTHS = (380, 100, 90, 120, 240, 140, 120, 90, 90)

    tree_imp.configure(columns=IMP_COLS, show="headings", selectmode="extended")
    _IMP_NUMERIC = {"quantity", "speed"}
    for key, hdr, w in zip(IMP_COLS, IMP_HEADERS, IMP_WIDTHS):
        tree_imp.heading(key, text=hdr)
        tree_imp.column(key, width=w, anchor=("e" if key in _IMP_NUMERIC else "w"))
    _enable_tree_sort(tree_imp)
    tree_imp.tag_configure("cat_exact",   background="#eaffea")
    tree_imp.tag_configure("cat_partial", background="#fff9d6")
    tree_imp.tag_configure("cat_missing", background="#ffecec")

    _iid_ix: Dict[str,int] = {}
    parsed_rows: List[Dict[str,Any]] = []
    _edI: Optional[tk.Entry]=None; _edItem: Optional[str]=None; _edCol: Optional[str]=None
    def _bbox_imp(item,col):
        try: b=tree_imp.bbox(item,col); return b if b else None
        except Exception: return None

    def _apply_cat_tag(iid: str, status: str):
        vals = list(tree_imp.item(iid,"values"))
        idx = IMP_COLS.index("cat_match")
        vals[idx] = "✓" if status=="exact" else ("≈" if status=="partial" else "—")
        tag = "cat_exact" if status=="exact" else ("cat_partial" if status=="partial" else "cat_missing")
        tree_imp.item(iid, values=tuple(vals), tags=(tag,))

    def _start_edit_imp(e=None):
        nonlocal _edI,_edItem,_edCol
        if tree_imp.identify("region", e.x, e.y)!="cell": return
        col = tree_imp.identify_column(e.x); row = tree_imp.identify_row(e.y)
        if not row or not col: return
        bx=_bbox_imp(row,col); 
        if not bx: return
        x,y,w,h = bx; col_idx=int(col[1:])-1; col_name=tree_imp["columns"][col_idx]
        cur = tree_imp.set(row,col_name)
        _edItem,_edCol=row,col
        _edI=tk.Entry(tree_imp); _edI.insert(0,cur); _edI.select_range(0,"end")
        _edI.focus_set(); _edI.place(x=x,y=y,width=w,height=h)
        def _commit(_e=None):
            val = _edI.get()
            if col_name in ("quantity",): val=_coerce_numeric(val)
            if col_name in ("volume",):   val=_normalize_volume_to_str(val)
            if col_name in ("flavor",):   val=_cleanup_flavor(val)
            tree_imp.set(_edItem,col_name,val)
            ix=_iid_ix.get(_edItem)
            if ix is not None and ix < len(parsed_rows):
                # Обновляем значение в parsed_rows (кроме cat_match - служебное поле)
                if col_name != "cat_match":
                    parsed_rows[ix][col_name] = val
                if col_name in ("name","volume"):
                    st=_catalog_match_status(tree_imp.set(_edItem,"name"), tree_imp.set(_edItem,"volume"))
                    _apply_cat_tag(_edItem, st)
            _cancel()
        def _cancel(_e=None):
            nonlocal _edI,_edItem,_edCol
            if _edI: _edI.destroy()
            _edI=_edItem=_edCol=None
        _edI.bind("<Return>",_commit); _edI.bind("<Escape>",_cancel); _edI.bind("<FocusOut>",_commit)
    tree_imp.bind("<Double-1>", _start_edit_imp); tree_imp.bind("<Return>", _start_edit_imp)

    def _import_from_clipboard():
        try: s = tab_import.clipboard_get()
        except Exception:
            messagebox.showwarning("Буфер обмена","Буфер обмена пуст или недоступен."); return
        if not s.strip():
            messagebox.showinfo("Буфер обмена","В буфере пусто."); return
        txt.delete("1.0","end"); txt.insert("1.0", s)
        lbl_info.config(text=f"Вставлено из буфера: {len(s)} символов")

    def _clear_input():
        txt.delete("1.0","end"); tree_imp.delete(*tree_imp.get_children(""))
        parsed_rows.clear(); _iid_ix.clear(); lbl_info.config(text="")

    def _run_parse():
        src = txt.get("1.0","end").strip()
        tree_imp.delete(*tree_imp.get_children("")); parsed_rows.clear(); _iid_ix.clear()
        if not src:
            lbl_info.config(text="Нет текста для распознавания"); return
        rows, profile = parse_clipboard_text(src)
        if not rows:
            lbl_info.config(text=f"Ничего не распознано (профиль: {profile})"); return

        dropped = 0
        for r in rows:
            name_src = r.get("name","") or ""
            # qty / volume из имени при необходимости
            if not r.get("quantity"):
                mq=_QTY_RX.search(name_src)
                if mq: r["quantity"]=_coerce_numeric(mq.group(1))
            if not r.get("volume"):
                mv=_VOL_TOKEN_RX.search(name_src)
                if mv: r["volume"]=_normalize_volume_to_str(mv.group(0))

            # извлечение типа/вкуса/бренда
            pp_type=pp_flavor=pp_brand=""
            if _pparse:
                try:
                    name_for_pp = _QTY_TAIL_RX.sub("", name_src)
                    name_for_pp = _VOL_TOKEN_RX.sub("", name_for_pp)
                    name_for_pp = re.sub(r"\s{2,}"," ", name_for_pp).strip(" ,;:-—")
                    pp = _pparse(name_for_pp, r.get("volume",""))
                    pp_type   = (pp.get("type") or "").capitalize()
                    pp_flavor = _cleanup_flavor(pp.get("flavor") or "")
                    pp_brand  = pp.get("brand") or ""
                except Exception:
                    pass
            if not pp_flavor or not pp_type or not pp_brand:
                t2,f2,b2 = _extract_type_flavor_brand(name_src, r.get("volume",""))
                pp_type = pp_type or t2
                pp_flavor = pp_flavor or f2
                pp_brand  = pp_brand or b2
            
            # Сохраняем распознанные type/flavor/brand в словарь
            r["type"] = pp_type
            r["flavor"] = pp_flavor
            r["brand"] = pp_brand
            
            # Убедимся что все поля из COL_KEYS присутствуют
            normalized_row = {k: r.get(k, "") for k in COL_KEYS}
            normalized_row["status"] = normalized_row.get("status") or "Planned"

            # Формируем values для отображения согласно IMP_COLS
            vals = [normalized_row.get(k, "") for k in IMP_COLS[:-1]] + [""]  # cat_match заполним позже
            iid = tree_imp.insert("", "end", values=tuple(vals))
            parsed_rows.append(normalized_row)
            _iid_ix[iid] = len(parsed_rows)-1
            _apply_cat_tag(iid, _catalog_match_status(normalized_row.get("name",""), normalized_row.get("volume","")))

        cat_state = "ON" if os.path.isfile(_catalog_path()) else "OFF"
        lbl_info.config(text=f"Распознано: {len(parsed_rows)} (профиль: {profile}; product_parse={'ON' if _pparse else 'OFF'}; catalog={cat_state})")

    def _enrich_preview_with_catalog():
        items = tree_imp.get_children("")
        if not items: return
        changed=0
        preserve_line = bool(var_preserve_line.get()); overwrite_speed = bool(var_overwrite_speed.get())
        for iid in items:
            vals = list(tree_imp.item(iid,"values"))
            ix = _iid_ix.get(iid)
            if ix is None or ix >= len(parsed_rows):
                continue
            
            row = parsed_rows[ix]
            enriched = _enrich_from_catalog(row, preserve_line_if_set=preserve_line,
                                            overwrite_speed_if_stronger=overwrite_speed)
            if enriched != row:
                changed += 1
                parsed_rows[ix] = enriched
                # Обновляем отображение
                new_vals = [enriched.get(k, "") for k in IMP_COLS[:-1]] + [vals[-1]]  # сохраняем cat_match
                tree_imp.item(iid, values=tuple(new_vals))
            _apply_cat_tag(iid, _catalog_match_status(enriched.get("name",""), enriched.get("volume","")))
        lbl_info.config(text=f"Обогащено: {changed}")

    def _show_mismatches():
        """Показать окно с продуктами, не найденными или частично найденными в каталоге"""
        items = tree_imp.get_children("")
        if not items:
            messagebox.showinfo("Несовпадения", "Нет распознанных строк.")
            return
        
        # Собираем несовпавшие и частично совпавшие
        _load_catalog_maps()
        results = []
        
        for iid in items:
            vals = list(tree_imp.item(iid, "values"))
            if len(vals) <= len(IMP_COLS) - 1:
                continue
                
            cat_match = vals[IMP_COLS.index("cat_match")]
            name = vals[IMP_COLS.index("name")]
            volume = vals[IMP_COLS.index("volume")]
            line = vals[IMP_COLS.index("line")]
            speed = vals[IMP_COLS.index("speed")]
            
            if cat_match == "—":
                # Полное несовпадение - новый продукт
                results.append({
                    "status": "Новый",
                    "import_name": name,
                    "import_volume": volume,
                    "import_line": line,
                    "import_speed": speed,
                    "catalog_name": "—",
                    "catalog_volume": "—",
                    "catalog_line": "—",
                    "catalog_speed": "—",
                })
            elif cat_match == "≈":
                # Частичное совпадение - найдем что в каталоге
                norm_name = _norm_name_match(name)
                cat_rec = _catalog_by_name.get(norm_name)
                
                if cat_rec:
                    results.append({
                        "status": "Конфликт",
                        "import_name": name,
                        "import_volume": volume,
                        "import_line": line,
                        "import_speed": speed,
                        "catalog_name": cat_rec.get("name", "—"),
                        "catalog_volume": cat_rec.get("container", "—"),
                        "catalog_line": cat_rec.get("line", "—"),
                        "catalog_speed": str(cat_rec.get("speed", "—")) if cat_rec.get("speed") else "—",
                    })
                else:
                    # Нашли по product_parse, но точной записи нет
                    results.append({
                        "status": "Новый",
                        "import_name": name,
                        "import_volume": volume,
                        "import_line": line,
                        "import_speed": speed,
                        "catalog_name": "Похожий есть",
                        "catalog_volume": "—",
                        "catalog_line": "—",
                        "catalog_speed": "—",
                    })
        
        if not results:
            messagebox.showinfo("Несовпадения", "✅ Все продукты точно найдены в каталоге!")
            return
        
        # Создаем окно с результатами
        win = tk.Toplevel(tab_import)
        win.title(f"Управление каталогом ({len(results)} шт.)")
        win.geometry("1400x700")
        
        ttk.Label(win, text="Сопоставление с каталогом:", font=("", 10, "bold")).pack(padx=10, pady=10, anchor="w")
        
        # Фрейм с таблицей
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        cols = ("action", "status", "import_name", "import_volume", "import_line", "import_speed",
                "catalog_name", "catalog_volume", "catalog_line", "catalog_speed")
        tree_miss = ttk.Treeview(frame, columns=cols, show="tree headings", selectmode="extended")
        
        tree_miss.heading("#0", text="☑")
        tree_miss.heading("action", text="Действие")
        tree_miss.heading("status", text="Статус")
        tree_miss.heading("import_name", text="Импорт: Имя")
        tree_miss.heading("import_volume", text="Импорт: Объём")
        tree_miss.heading("import_line", text="Импорт: Линия")
        tree_miss.heading("import_speed", text="Импорт: Скорость")
        tree_miss.heading("catalog_name", text="Каталог: Имя")
        tree_miss.heading("catalog_volume", text="Каталог: Объём")
        tree_miss.heading("catalog_line", text="Каталог: Линия")
        tree_miss.heading("catalog_speed", text="Каталог: Скорость")
        
        tree_miss.column("#0", width=30, stretch=False)
        tree_miss.column("action", width=120)
        tree_miss.column("status", width=80)
        tree_miss.column("import_name", width=250)
        tree_miss.column("import_volume", width=90)
        tree_miss.column("import_line", width=90)
        tree_miss.column("import_speed", width=90)
        tree_miss.column("catalog_name", width=250)
        tree_miss.column("catalog_volume", width=90)
        tree_miss.column("catalog_line", width=90)
        tree_miss.column("catalog_speed", width=90)
        
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree_miss.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree_miss.xview)
        tree_miss.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree_miss.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        
        tree_miss.tag_configure("new", background="#e8f4ff")
        tree_miss.tag_configure("conflict", background="#fff9d6")
        
        # Заполняем данные
        item_data = {}  # iid -> dict с данными
        for r in results:
            default_action = "Добавить новый" if r["status"] == "Новый" else "Обновить"
            tag = "new" if r["status"] == "Новый" else "conflict"
            
            vals = (default_action, r["status"], 
                   r["import_name"], r["import_volume"], r["import_line"], r["import_speed"],
                   r["catalog_name"], r["catalog_volume"], r["catalog_line"], r["catalog_speed"])
            iid = tree_miss.insert("", "end", text="☑", values=vals, tags=(tag,))
            item_data[iid] = r
        
        # Обработка клика по действию (переключение)
        def _toggle_action(event):
            region = tree_miss.identify("region", event.x, event.y)
            if region != "cell":
                return
            col = tree_miss.identify_column(event.x)
            row = tree_miss.identify_row(event.y)
            if not row or col != "#1":  # #1 = первая колонка (action)
                return
            
            vals = list(tree_miss.item(row, "values"))
            current_action = vals[0]
            r = item_data.get(row)
            if not r:
                return
            
            # Переключаем действие
            if r["status"] == "Новый":
                vals[0] = "Пропустить" if current_action == "Добавить новый" else "Добавить новый"
            else:  # Конфликт
                if current_action == "Обновить":
                    vals[0] = "Добавить новый"
                elif current_action == "Добавить новый":
                    vals[0] = "Пропустить"
                else:
                    vals[0] = "Обновить"
            
            tree_miss.item(row, values=vals)
        
        tree_miss.bind("<Double-1>", _toggle_action)
        
        # Кнопки управления
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        def _apply_changes():
            """Применить выбранные действия"""
            updates = []  # Для обновления существующих
            additions = []  # Для добавления новых
            
            for iid in tree_miss.get_children():
                vals = tree_miss.item(iid, "values")
                action = vals[0]
                if action == "Пропустить":
                    continue
                
                r = item_data[iid]
                entry = {
                    "name": r["import_name"],
                    "container": r["import_volume"],
                    "line": r["import_line"],
                    "speed": int(r["import_speed"]) if str(r["import_speed"]).strip().isdigit() else None,
                    "limit": None,
                    "action": ""
                }
                
                if action == "Обновить":
                    updates.append(entry)
                elif action == "Добавить новый":
                    additions.append(entry)
            
            # Применяем изменения
            all_entries = updates + additions
            if all_entries:
                added = _append_to_catalog_by_name(all_entries)
                messagebox.showinfo("Каталог", 
                    f"Готово!\n\nДобавлено новых: {added}\nОбновлено: {len(updates)}")
                
                # Обновляем подсветку в основной таблице
                for iid in tree_imp.get_children(""):
                    vals = list(tree_imp.item(iid, "values"))
                    _apply_cat_tag(iid, _catalog_match_status(
                        vals[IMP_COLS.index("name")], 
                        vals[IMP_COLS.index("volume")]
                    ))
                
                win.destroy()
            else:
                messagebox.showinfo("Каталог", "Нет действий для выполнения.")
        
        ttk.Label(btn_frame, text="Двойной клик по 'Действие' для изменения", 
                 foreground="#666", font=("", 9, "italic")).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Отмена", command=win.destroy).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="✓ Применить", command=_apply_changes, 
                  style="Accent.TButton").pack(side="right", padx=5)
    
    def _bulk_set_line(selected_only: bool):
        """Массово задать значение поля 'line' для выбранных или всех распознанных строк."""
        items = tree_imp.selection() if selected_only else tree_imp.get_children("")
        if not items:
            messagebox.showinfo("Импорт", "Нет подходящих строк для изменения.")
            return
        try:
            from tkinter import simpledialog
        except Exception:
            simpledialog = None  # type: ignore
        if simpledialog is None:
            messagebox.showerror("Импорт", "Не удалось открыть окно ввода.")
            return
        line_val = simpledialog.askstring(
            "Массовая замена линии",
            "Введите значение поля 'Линия' (например: Линия 3 или 3):",
            parent=tab_import
        )
        if line_val is None:
            return
        line_val = str(line_val).strip()
        if not line_val:
            messagebox.showinfo("Импорт", "Пустое значение линии не применяется.")
            return
        cnt = 0
        col_idx = IMP_COLS.index("line")
        for iid in items:
            vals = list(tree_imp.item(iid, "values"))
            if not vals:
                continue
            if col_idx < len(vals):
                vals[col_idx] = line_val
                tree_imp.item(iid, values=tuple(vals))
            ix = _iid_ix.get(iid)
            if ix is not None and ix < len(parsed_rows):
                parsed_rows[ix]["line"] = line_val
            cnt += 1
        lbl_info.config(text=f"Линия обновлена у {cnt} строк")

    def _bulk_set_speed(selected_only: bool):
        """Массово задать значение поля 'speed' (число) для выбранных или всех распознанных строк."""
        items = tree_imp.selection() if selected_only else tree_imp.get_children("")
        if not items:
            messagebox.showinfo("Импорт", "Нет подходящих строк для изменения.")
            return
        try:
            from tkinter import simpledialog
        except Exception:
            simpledialog = None  # type: ignore
        if simpledialog is None:
            messagebox.showerror("Импорт", "Не удалось открыть окно ввода.")
            return
        speed_raw = simpledialog.askstring(
            "Массовая замена скорости",
            "Введите скорость (шт/час):",
            parent=tab_import
        )
        if speed_raw is None:
            return
        speed_val = _coerce_numeric(str(speed_raw))
        if not speed_val:
            messagebox.showinfo("Импорт", "Некорректное значение скорости.")
            return
        cnt = 0
        col_idx = IMP_COLS.index("speed")
        for iid in items:
            vals = list(tree_imp.item(iid, "values"))
            if not vals:
                continue
            if col_idx < len(vals):
                vals[col_idx] = speed_val
                tree_imp.item(iid, values=tuple(vals))
            ix = _iid_ix.get(iid)
            if ix is not None and ix < len(parsed_rows):
                parsed_rows[ix]["speed"] = speed_val
            cnt += 1
        lbl_info.config(text=f"Скорость обновлена у {cnt} строк")

    def _add_selected_to_catalog():
        sels = tree_imp.selection() or tree_imp.get_children("")
        if not sels:
            messagebox.showinfo("Каталог","Нет выбранных строк."); return
        additions=[]
        for iid in sels:
            vals=list(tree_imp.item(iid,"values"))
            row={k: vals[i] if i<len(IMP_COLS) else "" for i,k in enumerate(IMP_COLS)}
            name=row.get("name","").strip()
            if not name: continue
            container = row.get("volume","").strip()
            speed=row.get("speed",""); speed_val=None
            if str(speed).strip().isdigit(): speed_val=int(speed)
            line=row.get("line","") or ""
            additions.append({
                "name":name,"line":line,"container":container,
                "speed":speed_val,"limit":None,"action":""
            })
        if not additions:
            messagebox.showinfo("Каталог","Нечего добавлять."); return
        added = _append_to_catalog_by_name(additions)
        for iid in sels:
            vals=list(tree_imp.item(iid,"values"))
            _apply_cat_tag(iid, _catalog_match_status(vals[IMP_COLS.index("name")], vals[IMP_COLS.index("volume")]))
        messagebox.showinfo("Каталог", f"Добавлено новых: {added}. Остальные — обновлены пустые поля.")

    def _apply_to_plan():
        """Перенос распознанных строк в План.
        Если у строки пустой JobID — присваиваем новый (J-YYMMDD-LNN-XXX)."""
        if not parsed_rows:
            messagebox.showinfo("Импорт", "Нет распознанных строк.")
            return

        preserve_line = bool(var_preserve_line.get())
        overwrite_speed = bool(var_overwrite_speed.get())

        # Синхронизируем значения из UI с parsed_rows (на случай ручных правок)
        for iid in tree_imp.get_children(""):
            vals = tree_imp.item(iid, "values")
            ix = _iid_ix.get(iid)
            if ix is None or ix >= len(parsed_rows):
                continue
            # Обновляем все поля кроме cat_match (последний)
            for i, col_key in enumerate(IMP_COLS[:-1]):
                if i < len(vals):
                    parsed_rows[ix][col_key] = vals[i]

        # уже занятые JobID в Плане
        existing_ids = _collect_existing_job_ids(tree)

        added = 0
        # Сохраняем порядок предпросмотра: добавляем сверху вниз
        for r in parsed_rows:
            # обогащение из каталога (как было)
            r = _enrich_from_catalog(
                r,
                preserve_line_if_set=preserve_line,
                overwrite_speed_if_stronger=overwrite_speed
            )

            # присвоить JobID, если пустой
            if not str(r.get("job_id", "")).strip():
                r["job_id"] = _next_job_id(existing_ids, r.get("line", ""))

            _insert(r)
            added += 1

        info_lbl.config(text=f"Импортировано: {added} строк (JobID назначен где пусто)")
        sub.select(tab_plan)


    btn_clip.configure(command=_import_from_clipboard)
    btn_clear.configure(command=_clear_input)
    btn_parse.configure(command=_run_parse)
    btn_enrich.configure(command=_enrich_preview_with_catalog)
    btn_show_miss.configure(command=_show_mismatches)
    btn_addcat.configure(command=_add_selected_to_catalog)
    btn_set_line_sel.configure(command=lambda: _bulk_set_line(True))
    btn_set_line_all.configure(command=lambda: _bulk_set_line(False))
    btn_set_speed_sel.configure(command=lambda: _bulk_set_speed(True))
    btn_set_speed_all.configure(command=lambda: _bulk_set_speed(False))
    btn_apply.configure(command=_apply_to_plan)

    tab_import.bind_all("<Control-Return>", lambda e: (_apply_to_plan(), "break"))
    tab_import.bind_all("<Control-v>", lambda e: (_import_from_clipboard(), "break"))

    # экспорт наружу (если где-то используется)
    tab_planning.tree_plan   = tree
    tab_planning.save_json   = _save_json
    tab_planning.load_json   = _load_json
    tab_planning.tree_import = tree_imp
    tab_planning.parse_text  = _run_parse
    tab_planning.apply_import= _apply_to_plan
    tab_planning.input_text  = txt

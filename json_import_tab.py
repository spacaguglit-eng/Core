# -*- coding: utf-8 -*-
"""
json_import_tab.py — вкладка «Импорт JSON» (OEE-таблица) + автообновление факта
-------------------------------------------------------------------------------
• Кнопка «Открыть JSON…» (один раз указать путь)
• Путь сохраняется в settings_oee.json рядом с модулем
• Фоновый тихий мониторинг файла (mtime): при изменении — перезагрузка JSON
  и авто-подтяжка «Факт, шт» в Плане по совпадающему job_id
• Без всплывающих окон (кроме явных ошибок чтения при ручном выборе файла)
• Таблица в этой вкладке обновляется «для вида», но без диалогов
"""

from __future__ import annotations
import json, os, math, time, re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any, Optional

# ---------------------------------------------------------------------
_THIS_DIR = os.path.dirname(__file__)
_SETTINGS_PATH = os.path.join(_THIS_DIR, "settings_oee.json")

HEADERS = [
    "Job ID","Продукт","Линия","День","Смена","Начало","Конец","Длит (мин)",
    "Σ простой (мин)","% простоя","Событий","План. простой (мин)",
    "EffMin (мин)","Ном. скорость (ш)","Потолок (шт)","Факт (шт)","OEE, %"
]

def _load_settings() -> dict:
    try:
        if os.path.isfile(_SETTINGS_PATH):
            with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
                return d if isinstance(d, dict) else {}
    except Exception:
        pass
    return {}

def _save_settings(d: dict) -> None:
    try:
        with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _flatten_payload(payload: Any) -> List[Dict[str, Any]]:
    if isinstance(payload, list):
        return [r for r in payload if isinstance(r, dict)]
    if isinstance(payload, dict):
        if "data" in payload and isinstance(payload["data"], list):
            return payload["data"]
        for v in payload.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return v
    return []

def _num(x) -> float:
    if x in (None, ""): return math.nan
    s = str(x).replace(" ", "").replace(",", ".")
    try: return float(s)
    except Exception: return math.nan

def _fmt(x, nd=0):
    if isinstance(x, float) and (math.isnan(x) or math.isinf(x)): return ""
    if x is None: return ""
    if nd == 0: return str(int(round(x)))
    return f"{x:.{nd}f}".rstrip("0").rstrip(".")

def _shift_from_time(t: str) -> str:
    if not t or ":" not in t: return ""
    try:
        h = int(t.split(":")[0])
        return "День" if 8 <= h < 20 else "Ночь"
    except Exception:
        return ""

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

# ---------------------------------------------------------------------
class JsonImportTab:
    def __init__(self, nb: ttk.Notebook, on_import=None, parent_frame=None):
        self._nb = nb
        self._on_import = on_import  # необязательный колбэк: on_import(kind, block_name, headers, rows, meta)
        
        # Если передан parent_frame, используем его, иначе создаем новую вкладку
        if parent_frame:
            self._tab = parent_frame
        else:
            self._tab = ttk.Frame(nb)
            nb.add(self._tab, text="Импорт JSON")

        self._rows: List[List[Any]] = []
        self._all_records: List[Dict[str, Any]] = []  # Все записи для фильтрации
        self._json_path: Optional[str] = None
        self._last_mtime: Optional[float] = None
        self._watch_period_ms = 3000  # проверяем раз в 3 секунды
        
        # Переменные для сортировки
        self.sort_column: Optional[int] = None
        self.sort_reverse: bool = False

        # Создаем основной контейнер с прокруткой
        self._main_container = ttk.Frame(self._tab)
        self._main_container.pack(fill="both", expand=True)
        
        # Строим все компоненты интерфейса
        self._build_header(self._main_container)
        self._build_statistics_panel(self._main_container)
        self._build_controls_panel(self._main_container)
        self._build_table(self._main_container)
        self._build_status_bar(self._main_container)

        # восстановим путь из настроек и запустим тихую подгрузку/мониторинг
        st = _load_settings()
        path = st.get("oee_json_path", "")
        if path and os.path.isfile(path):
            self._set_path_and_start(path, initial_load=True, silent=True)

    # ---------- Шапка с управлением файлом ----------
    def _build_header(self, parent):
        """Верхняя панель с выбором файла"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill="x", padx=8, pady=(8, 0))
        
        # Левая часть - выбор файла
        left_section = ttk.Frame(header_frame)
        left_section.pack(side="left", fill="x", expand=True)
        
        btn_open = ttk.Button(left_section, text="📂 Открыть JSON файл", 
                             command=self._open_json, width=20)
        btn_open.pack(side="left", padx=(0, 12))
        
        # Информация о файле
        file_info_frame = ttk.Frame(left_section)
        file_info_frame.pack(side="left", fill="x", expand=True)
        
        ttk.Label(file_info_frame, text="Файл:", foreground="#666").pack(side="left")
        self.lbl_file = ttk.Label(file_info_frame, text="не выбран", 
                                  foreground="#333", font=("TkDefaultFont", 9))
        self.lbl_file.pack(side="left", padx=(6, 0))
        
        # Правая часть - статус обновления
        self.lbl_status = ttk.Label(header_frame, text="● Готов", 
                                    foreground="#28a745", font=("TkDefaultFont", 9))
        self.lbl_status.pack(side="right", padx=(8, 0))
    
    # ---------- Панель статистики с карточками ----------
    def _build_statistics_panel(self, parent):
        """Панель с ключевыми метриками в виде карточек"""
        stats_container = ttk.Frame(parent)
        stats_container.pack(fill="x", padx=8, pady=(8, 0))
        
        # Создаем карточки для метрик
        cards_frame = ttk.Frame(stats_container)
        cards_frame.pack(fill="x")
        
        # Карточка 1: Записи
        self.card_records = self._create_stat_card(cards_frame, "📊 Записей", "0", "#007bff")
        self.card_records.pack(side="left", fill="x", expand=True, padx=(0, 6))
        
        # Карточка 2: OEE
        self.card_oee = self._create_stat_card(cards_frame, "📈 Средний OEE", "— %", "#28a745")
        self.card_oee.pack(side="left", fill="x", expand=True, padx=(0, 6))
        
        # Карточка 3: Простои
        self.card_downtimes = self._create_stat_card(cards_frame, "⚠️ Простоев", "0", "#ffc107")
        self.card_downtimes.pack(side="left", fill="x", expand=True, padx=(0, 6))
        
        # Карточка 4: Общее время простоев
        self.card_downtime_min = self._create_stat_card(cards_frame, "⏱️ Время простоев", "0 мин", "#dc3545")
        self.card_downtime_min.pack(side="left", fill="x", expand=True)
    
    def _create_stat_card(self, parent, title, value, color):
        """Создает карточку метрики"""
        card = ttk.LabelFrame(parent, padding=12, relief="flat")
        
        # Заголовок
        title_label = ttk.Label(card, text=title, foreground="#666", 
                               font=("TkDefaultFont", 9))
        title_label.pack(anchor="w", pady=(0, 4))
        
        # Значение
        value_label = ttk.Label(card, text=value, foreground=color, 
                               font=("TkDefaultFont", 16, "bold"))
        value_label.pack(anchor="w")
        
        # Сохраняем ссылку на label значения для обновления
        card.value_label = value_label
        card.title_label = title_label
        
        return card
    
    # ---------- Панель управления и фильтров ----------
    def _build_controls_panel(self, parent):
        """Панель с фильтрами и настройками"""
        controls_container = ttk.LabelFrame(parent, text="Управление данными", padding=12)
        controls_container.pack(fill="x", padx=8, pady=(8, 0))
        
        # Верхняя строка - фильтры
        filters_row = ttk.Frame(controls_container)
        filters_row.pack(fill="x", pady=(0, 8))
        
        # Фильтр по линии
        line_group = ttk.Frame(filters_row)
        line_group.pack(side="left", padx=(0, 16))
        ttk.Label(line_group, text="Линия:", font=("TkDefaultFont", 9)).pack(side="left", padx=(0, 6))
        self.line_filter = ttk.Combobox(line_group, width=18, state="readonly", 
                                       font=("TkDefaultFont", 9))
        self.line_filter.pack(side="left")
        self.line_filter.bind("<<ComboboxSelected>>", self._apply_filters)
        
        # Фильтр по дню
        day_group = ttk.Frame(filters_row)
        day_group.pack(side="left", padx=(0, 16))
        ttk.Label(day_group, text="День:", font=("TkDefaultFont", 9)).pack(side="left", padx=(0, 6))
        self.day_filter = ttk.Combobox(day_group, width=18, state="readonly", 
                                      font=("TkDefaultFont", 9))
        self.day_filter.pack(side="left")
        self.day_filter.bind("<<ComboboxSelected>>", self._apply_filters)
        
        # Фильтр по тексту (поиск)
        search_group = ttk.Frame(filters_row)
        search_group.pack(side="left", padx=(0, 16))
        ttk.Label(search_group, text="Поиск:", font=("TkDefaultFont", 9)).pack(side="left", padx=(0, 6))
        self.search_entry = ttk.Entry(search_group, width=20, font=("TkDefaultFont", 9))
        self.search_entry.pack(side="left")
        self.search_entry.bind("<KeyRelease>", lambda e: self._apply_filters())
        # Иконка поиска
        ttk.Label(search_group, text="🔍", font=("TkDefaultFont", 10)).pack(side="left", padx=(4, 0))
        
        # Кнопки управления
        buttons_row = ttk.Frame(controls_container)
        buttons_row.pack(side="right")
        
        btn_reset = ttk.Button(buttons_row, text="🔄 Сбросить", 
                              command=self._reset_filters, width=15)
        btn_reset.pack(side="left", padx=(0, 8))
        
        # Кнопка экспорта
        btn_export = ttk.Button(buttons_row, text="💾 Экспорт", 
                                command=self._export_data, width=15)
        btn_export.pack(side="left", padx=(0, 8))
        
        # Чекбокс показа простоев
        self.show_downtimes_var = tk.BooleanVar(value=True)
        chk_downtimes = ttk.Checkbutton(buttons_row, 
                                       text="📋 Показать простои", 
                                       variable=self.show_downtimes_var,
                                       command=self._toggle_downtimes)
        chk_downtimes.pack(side="left")
        
        # Заполнитель для выравнивания
        ttk.Frame(filters_row).pack(side="left", fill="x", expand=True)

    # ---------- Таблица данных ----------
    def _build_table(self, parent):
        """Основная таблица с данными OEE"""
        table_wrapper = ttk.LabelFrame(parent, text="📋 Данные производства", padding=8)
        table_wrapper.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        # Контейнер для таблицы и скроллбаров
        table_container = ttk.Frame(table_wrapper)
        table_container.pack(fill="both", expand=True)
        
        # Таблица с улучшенным видом
        self.tree = ttk.Treeview(table_container, show="tree headings", height=22,
                                style="Custom.Treeview")
        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_container.rowconfigure(0, weight=1)
        table_container.columnconfigure(0, weight=1)
        
        # Настройка стиля таблицы с улучшенным оформлением
        style = ttk.Style()
        style.configure("Custom.Treeview", rowheight=26, font=("Segoe UI", 9))
        style.configure("Custom.Treeview.Heading", font=("Segoe UI", 9, "bold"), 
                       background="#f0f0f0", foreground="#333")
        style.map("Custom.Treeview.Heading", 
                 background=[("active", "#e0e0e0"), ("pressed", "#d0d0d0")])

        # Настройка колонок
        self.tree["columns"] = [f"c{i}" for i in range(len(HEADERS))]
        self.tree.column("#0", width=25, stretch=False, minwidth=25)  # Колонка для иконок раскрытия
        
        column_widths = {
            "Job ID": 110, "Продукт": 320, "Линия": 110, "День": 110, "Смена": 85,
            "Начало": 125, "Конец": 125, "Длит (мин)": 95, "Σ простой (мин)": 115,
            "% простоя": 95, "Событий": 85, "План. простой (мин)": 115,
            "EffMin (мин)": 95, "Ном. скорость (ш)": 125, "Потолок (шт)": 105,
            "Факт (шт)": 105, "OEE, %": 85
        }
        
        for i, h in enumerate(HEADERS):
            anchor = "e" if h not in ("Job ID","Продукт","Линия","День","Смена","Начало","Конец") else "w"
            # Добавляем индикатор сортировки в заголовок
            heading_text = h
            self.tree.heading(f"c{i}", text=heading_text, anchor=anchor,
                            command=lambda c=i, col=h: self._sort_by_column(c, col))
            width = column_widths.get(h, 110 if anchor=="e" else 140)
            self.tree.column(f"c{i}", width=width, anchor=anchor, minwidth=60)
        
        # Улучшенная цветовая индикация через теги
        self.tree.tag_configure("high_oee", background="#e8f5e9", foreground="#155724")  # Зеленый для высокого OEE
        self.tree.tag_configure("low_oee", background="#fff3e0", foreground="#856404")  # Желтый для низкого OEE
        self.tree.tag_configure("very_low_oee", background="#ffebee", foreground="#721c24")  # Красный для очень низкого OEE
        self.tree.tag_configure("high_downtime", background="#fce4ec", foreground="#721c24")  # Розовый для высоких простоев
        self.tree.tag_configure("downtime_detail", background="#f5f5f5", foreground="#495057", 
                               font=("Segoe UI", 8))  # Серый для деталей простоев
        
        # Улучшенное чередование цветов для основных строк
        self.tree.tag_configure("row_even", background="#ffffff")
        self.tree.tag_configure("row_odd", background="#f8f9fa")
        
        # Выделение выбранной строки
        style.map("Custom.Treeview",
                 background=[("selected", "#007bff")],
                 foreground=[("selected", "white")])
        
        # Обработка двойного клика для раскрытия простоев
        self.tree.bind("<Double-1>", self._on_row_double_click)
        
        # Обработка клика для показа деталей простоя (для дочерних элементов)
        self.tree.bind("<Button-1>", self._on_row_click)
        
        # Хранилище данных о простоях для каждой строки
        self.downtimes_data = {}  # {row_id: [downtime_dict, ...]}
        self.expanded_rows = set()  # Множество раскрытых строк
        self._tooltip_window = None  # Всплывающее окно с подсказкой
        self._tooltip_item = None  # Текущий элемент с подсказкой
        self._tooltips = {}  # {item_id: tooltip_text} - хранилище подсказок
        self._sorted_data = []  # Отсортированные данные для быстрого доступа
    
    # ---------- Статусная строка ----------
    def _build_status_bar(self, parent):
        """Нижняя строка с дополнительной информацией"""
        status_bar = ttk.Frame(parent)
        status_bar.pack(fill="x", padx=8, pady=(4, 8))
        
        # Левая часть - подсказки
        tips_frame = ttk.Frame(status_bar)
        tips_frame.pack(side="left", fill="x", expand=True)
        
        ttk.Label(tips_frame, text="💡 Двойной клик — раскрыть/свернуть простои | ",
                 foreground="#666", font=("TkDefaultFont", 8)).pack(side="left")
        ttk.Label(tips_frame, text="Клик по заголовку колонки — сортировка | ",
                 foreground="#666", font=("TkDefaultFont", 8)).pack(side="left")
        
        # Правая часть - количество отображаемых записей и информация о сортировке
        info_frame = ttk.Frame(status_bar)
        info_frame.pack(side="right")
        
        self.lbl_sort_info = ttk.Label(info_frame, text="", 
                                       foreground="#28a745", 
                                       font=("TkDefaultFont", 8))
        self.lbl_sort_info.pack(side="left", padx=(0, 10))
        
        self.lbl_record_count = ttk.Label(info_frame, text="", 
                                          foreground="#007bff", 
                                          font=("TkDefaultFont", 8, "bold"))
        self.lbl_record_count.pack(side="left")

    # ---------- Открытие файла пользователем ----------
    def _open_json(self):
        path = filedialog.askopenfilename(
            title="Выбрать JSON",
            filetypes=[("JSON файлы", "*.json"), ("Все файлы", "*.*")]
        )
        if not path: return
        try:
            # сохранить путь и запустить мониторинг
            self._set_path_and_start(path, initial_load=True, silent=True)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать JSON:\n{e}")

    # ---------- Настройка пути + старт наблюдателя ----------
    def _set_path_and_start(self, path: str, initial_load: bool, silent: bool):
        self._json_path = path
        self.lbl_file.config(text=f"Файл: {os.path.basename(path)}")
        st = _load_settings()
        st["oee_json_path"] = path
        _save_settings(st)
        if initial_load:
            self._load_apply_json(silent=silent)
        # стартуем таймер наблюдения
        self._schedule_watch()

    # ---------- Сохранение плана в jobs_plan.json ----------
    def _save_plan_to_json(self):
        """Сохраняет текущее состояние плана в jobs_plan.json"""
        try:
            # найдём вкладку «Планирование»
            nb = self._nb
            tab_plan = None
            for tid in nb.tabs():
                if nb.tab(tid, "text") == "Планирование":
                    tab_plan = nb.nametowidget(tid)
                    break
            if not tab_plan or not hasattr(tab_plan, "tree_plan"):
                return

            # вызываем метод сохранения из planning_tab
            if hasattr(tab_plan, "_save_json"):
                tab_plan._save_json()
        except Exception as e:
            print(f"[ERROR] Ошибка сохранения плана: {e}")

    # ---------- План: обновление факта по карте job_id -> fact ----------
    def _apply_fact_to_plan(self, fact_map: Dict[str, int]) -> int:
        """Тихо обновляет fact_qty/прогресс/процент в Плане по job_id. Возвращает кол-во обновлённых строк."""
        try:
            # найдём вкладку «Планирование»
            nb = self._nb
            tab_plan = None
            for tid in nb.tabs():
                if nb.tab(tid, "text") == "Планирование":
                    tab_plan = nb.nametowidget(tid)
                    break
            if not tab_plan or not hasattr(tab_plan, "tree_plan"):
                return 0

            tree = tab_plan.tree_plan
            # используем правильную структуру из planning_tab.py
            if hasattr(tab_plan, "COL_KEYS"):
                col_keys = list(tab_plan.COL_KEYS)
            else:
                # правильный порядок из planning_tab.py
                col_keys = [
                    "priority","job_id","name","volume","flavor","brand","type",
                    "quantity","line","speed","speed_source","status","fact_qty","progress"
                ]

            qty_idx  = col_keys.index("quantity") if "quantity" in col_keys else 7
            fact_idx = col_keys.index("fact_qty") if "fact_qty" in col_keys else 12
            prog_idx = col_keys.index("progress") if "progress" in col_keys else 13
            perc_idx = col_keys.index("percent_done") if "percent_done" in col_keys else 13

            updated = 0
            # Ищем данные в группах линий (дочерние элементы групп)
            for group_id in tree.get_children(""):
                # Проверяем дочерние элементы группы
                for iid in tree.get_children(group_id):
                    vals = list(tree.item(iid, "values"))
                    if not vals: 
                        continue
                    job_id = str(vals[1])  # job_id теперь на позиции 1
                    if job_id in fact_map:
                        # plan qty
                        plan_qty = 0
                        try:
                            plan_qty = int(str(vals[qty_idx]).replace(" ", ""))
                        except Exception:
                            pass
                        fact_qty = int(fact_map[job_id])
                        vals[fact_idx] = str(fact_qty)
                        if plan_qty > 0:
                            pct = fact_qty / plan_qty * 100
                            vals[prog_idx] = f"{fact_qty} / {plan_qty}"
                            # выводим как в план-таблице: десятые и запятая не обязательны
                            vals[perc_idx] = f"{pct:.1f}%"
                        tree.item(iid, values=tuple(vals))
                        updated += 1
            return updated
        except Exception:
            return 0

    # ---------- Загрузка JSON и заполнение таблицы + применение факта ----------
    def _load_apply_json(self, silent: bool):
        if not self._json_path or not os.path.isfile(self._json_path):
            return
        try:
            mtime = os.path.getmtime(self._json_path)
            self._last_mtime = mtime
            with open(self._json_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            records = _flatten_payload(payload)
            if not records:
                # тихо очистим таблицу, если пусто
                self.tree.delete(*self.tree.get_children())
                self._rows.clear()
                return

            # заполняем таблицу (без окон)
            self.tree.delete(*self.tree.get_children())
            self._tooltips.clear()  # Очищаем подсказки
            rows_out: List[List[Any]] = []
            fact_map: Dict[str, int] = {}

            # Сохраняем все записи для фильтрации
            self._all_records = records
            
            # Обновляем фильтры
            self._update_filters()
            
            # Фильтруем записи
            filtered_records = self._filter_records(records)
            
            # Сортируем записи, если выбрана колонка
            if self.sort_column is not None:
                filtered_records = self._sort_records(filtered_records)
            
            for r in filtered_records:
                job_id = r.get("job_id", "")
                product = r.get("product", "")
                line = r.get("line", "")
                day = r.get("date", "")
                start = r.get("start", "")
                end = r.get("end", "")
                shift = _shift_from_time(start)
                dur = _num(r.get("duration_min", ""))
                
                # Обработка простоев
                downtimes = r.get("downtimes", []) if isinstance(r.get("downtimes"), list) else []
                sum_dt = 0.0
                plan_dt = _num(r.get("planned_downtime_min", 0))
                
                # Вычисляем сумму простоев
                for dt_item in downtimes:
                    if isinstance(dt_item, dict):
                        dt_duration = _num(dt_item.get("duration_min", dt_item.get("duration", 0)))
                        sum_dt += dt_duration
                
                events = len(downtimes)
                speed = _num(r.get("speed", ""))
                fact = _num(r.get("fact", ""))

                if job_id and not (isinstance(fact, float) and math.isnan(fact)):
                    try:
                        fact_map[job_id] = int(round(float(fact)))
                    except Exception:
                        pass

                pct_dt = (sum_dt / dur * 100) if (dur and dur > 0) else 0
                effmin = max(0, (dur or 0) - sum_dt - plan_dt)
                ceil_units = effmin * speed / 60 if speed and effmin else 0
                oee = (fact / ceil_units * 100) if (ceil_units and ceil_units > 0) else 0

                row = [
                    job_id, product, line, day, shift, start, end,
                    _fmt(dur), _fmt(sum_dt), _fmt(pct_dt, 1),
                    _fmt(events), _fmt(plan_dt), _fmt(effmin),
                    _fmt(speed), _fmt(ceil_units), _fmt(fact), _fmt(oee, 1)
                ]
                rows_out.append(row)
                
                # Определяем теги для цветовой индикации
                tags = []
                
                # Чередование цветов строк
                row_index = len(self.tree.get_children())
                if row_index % 2 == 0:
                    tags.append("row_even")
                else:
                    tags.append("row_odd")
                
                # OEE индикация
                if not math.isnan(oee) and oee > 0:
                    if oee >= 85:
                        tags.append("high_oee")
                    elif oee >= 70:
                        tags.append("low_oee")
                    else:
                        tags.append("very_low_oee")
                
                if pct_dt > 20:  # Высокий процент простоев
                    tags.append("high_downtime")
                
                # Сохраняем данные о простоях для этой строки
                row_id = f"{job_id}_{line}_{day}"
                self.downtimes_data[row_id] = downtimes
                
                # Вставляем строку с возможностью раскрытия, если есть простои
                if downtimes:
                    item_id = self.tree.insert("", "end", text="▶", values=row, tags=tuple(tags))
                else:
                    item_id = self.tree.insert("", "end", text="", values=row, tags=tuple(tags))
                
                # Если включено отображение простоев, добавляем их сразу
                if self.show_downtimes_var.get() and downtimes:
                    self._add_downtimes_to_tree(item_id, downtimes)
                    self.expanded_rows.add(item_id)
            
            # Обновляем счетчик записей в статусной строке
            if hasattr(self, 'lbl_record_count'):
                filtered_count = len(rows_out)
                total_count = len(records)
                self.lbl_record_count.config(
                    text=f"Показано: {filtered_count}" + 
                         (f" из {total_count}" if filtered_count < total_count else "")
                )
            
            # Обновляем информацию о сортировке
            if hasattr(self, 'lbl_sort_info') and self.sort_column is not None:
                col_name = HEADERS[self.sort_column]
                direction = "▼" if self.sort_reverse else "▲"
                self.lbl_sort_info.config(text=f"Сортировка: {col_name} {direction}")
            elif hasattr(self, 'lbl_sort_info'):
                self.lbl_sort_info.config(text="")

            self._rows = rows_out
            
            # Обновляем статистику
            self._update_statistics(filtered_records)

            # тихо применяем факт к Плану
            if fact_map:
                updated = self._apply_fact_to_plan(fact_map)
                if updated > 0:
                    # сохраняем обновленный план в jobs_plan.json
                    self._save_plan_to_json()
        except Exception as e:
            # никаких окон, просто молчим, но логируем ошибку
            print(f"[ERROR] Ошибка загрузки JSON: {e}")
            import traceback
            traceback.print_exc()

    # ---------- Мониторинг файла ----------
    def _schedule_watch(self):
        # чтобы не множить таймеры, можно просто перезаписывать — Tk сам отработает
        self._tab.after(self._watch_period_ms, self._watch_once)

    def _watch_once(self):
        try:
            if self._json_path and os.path.isfile(self._json_path):
                mtime = os.path.getmtime(self._json_path)
                if (self._last_mtime is None) or (mtime > (self._last_mtime or 0)):
                    # файл новый или обновился — подхватить тихо
                    self._load_apply_json(silent=True)
        finally:
            self._schedule_watch()
    
    # ---------- Фильтрация ----------
    def _update_filters(self):
        """Обновление списков фильтров"""
        if not self._all_records:
            return
        
        # Собираем уникальные значения
        lines = set()
        days = set()
        
        for r in self._all_records:
            line = r.get("line", "")
            day = r.get("date", "")
            if line:
                lines.add(line)
            if day:
                days.add(day)
        
        # Обновляем комбобоксы
        line_values = ["Все"] + sorted(list(lines))
        day_values = ["Все"] + sorted(list(days))
        
        self.line_filter["values"] = line_values
        self.day_filter["values"] = day_values
        
        # Устанавливаем значения по умолчанию, если они еще не установлены
        if not self.line_filter.get():
            self.line_filter.set("Все")
        if not self.day_filter.get():
            self.day_filter.set("Все")
        
        # Инициализируем поле поиска, если его еще нет
        if not hasattr(self, 'search_entry'):
            # Это может произойти при первой загрузке до создания интерфейса
            pass
    
    def _filter_records(self, records: List[Dict]) -> List[Dict]:
        """Фильтрация записей по выбранным фильтрам"""
        filtered = records
        
        # Фильтр по линии
        line_value = self.line_filter.get()
        if line_value and line_value != "Все":
            filtered = [r for r in filtered if r.get("line", "") == line_value]
        
        # Фильтр по дню
        day_value = self.day_filter.get()
        if day_value and day_value != "Все":
            filtered = [r for r in filtered if r.get("date", "") == day_value]
        
        # Фильтр по поисковому запросу
        search_text = self.search_entry.get().strip().lower() if hasattr(self, 'search_entry') else ""
        if search_text:
            filtered = [r for r in filtered if self._matches_search(r, search_text)]
        
        return filtered
    
    def _matches_search(self, record: Dict, search_text: str) -> bool:
        """Проверка, соответствует ли запись поисковому запросу"""
        # Ищем во всех текстовых полях записи
        fields_to_search = [
            str(record.get("job_id", "")),
            str(record.get("product", "")),
            str(record.get("line", "")),
            str(record.get("date", "")),
            str(record.get("start", "")),
            str(record.get("end", "")),
        ]
        # Проверяем каждый простои
        downtimes = record.get("downtimes", [])
        if isinstance(downtimes, list):
            for dt in downtimes:
                if isinstance(dt, dict):
                    fields_to_search.extend([
                        str(dt.get("category", "")),
                        str(dt.get("reason", "")),
                        str(dt.get("description", "")),
                    ])
        
        combined_text = " ".join(fields_to_search).lower()
        return search_text in combined_text
    
    def _calculate_oee_for_sort(self, r: Dict) -> float:
        """Вычисляет OEE для сортировки"""
        try:
            dur = _num(r.get("duration_min", 0))
            fact = _num(r.get("fact", 0))
            speed = _num(r.get("speed", 0))
            dts = r.get("downtimes", [])
            plan_dt = _num(r.get("planned_downtime_min", 0))
            
            sum_dt = sum(_num(dt.get("duration_min", dt.get("duration", 0))) 
                       for dt in (dts or []) if isinstance(dt, dict))
            
            effmin = max(0, dur - sum_dt - plan_dt)
            
            if dur > 0 and speed > 0 and effmin > 0:
                ceil_units = effmin * speed / 60
                if ceil_units > 0:
                    return (fact / ceil_units) * 100
            return 0.0
        except:
            return 0.0
    
    def _sort_records(self, records: List[Dict]) -> List[Dict]:
        """Сортировка записей по выбранной колонке"""
        if self.sort_column is None:
            return records
        
        def get_sort_key(r: Dict) -> Any:
            """Извлекает значение для сортировки из записи"""
            col_name = HEADERS[self.sort_column]
            
            # Маппинг колонок на поля записи
            field_map = {
                "Job ID": "job_id",
                "Продукт": "product",
                "Линия": "line",
                "День": "date",
                "Смена": lambda r: _shift_from_time(r.get("start", "")),
                "Начало": "start",
                "Конец": "end",
                "Длит (мин)": lambda r: _num(r.get("duration_min", 0)),
                "Σ простой (мин)": lambda r: sum(_num(dt.get("duration_min", dt.get("duration", 0))) 
                                                   for dt in (r.get("downtimes", []) or []) 
                                                   if isinstance(dt, dict)),
                "% простоя": lambda r: ((_num(sum(_num(dt.get("duration_min", dt.get("duration", 0))) 
                                                  for dt in (r.get("downtimes", []) or []) 
                                                  if isinstance(dt, dict))) / _num(r.get("duration_min", 1)) * 100) 
                                       if _num(r.get("duration_min", 0)) > 0 else 0),
                "Событий": lambda r: len(r.get("downtimes", []) or []),
                "План. простой (мин)": lambda r: _num(r.get("planned_downtime_min", 0)),
                "EffMin (мин)": lambda r: max(0, (_num(r.get("duration_min", 0)) or 0) - 
                                              sum(_num(dt.get("duration_min", dt.get("duration", 0))) 
                                                  for dt in (r.get("downtimes", []) or []) 
                                                  if isinstance(dt, dict)) - 
                                              _num(r.get("planned_downtime_min", 0))),
                "Ном. скорость (ш)": lambda r: _num(r.get("speed", 0)),
                "Потолок (шт)": lambda r: (max(0, (_num(r.get("duration_min", 0)) or 0) - 
                                                 sum(_num(dt.get("duration_min", dt.get("duration", 0))) 
                                                     for dt in (r.get("downtimes", []) or []) 
                                                     if isinstance(dt, dict)) - 
                                           _num(r.get("planned_downtime_min", 0))) * 
                                          _num(r.get("speed", 0)) / 60 
                                          if _num(r.get("speed", 0)) > 0 else 0),
                "Факт (шт)": lambda r: _num(r.get("fact", 0)),
                "OEE, %": lambda r: self._calculate_oee_for_sort(r)
            }
            
            if col_name in field_map:
                field = field_map[col_name]
                if callable(field):
                    try:
                        val = field(r)
                    except:
                        val = ""
                else:
                    val = r.get(field, "")
            else:
                val = ""
            
            # Преобразуем в тип для сортировки
            if isinstance(val, (int, float)):
                if math.isnan(val):
                    return (1, 0) if self.sort_reverse else (0, 0)
                return (0, val) if val >= 0 else (1, abs(val))
            val_str = str(val).lower()
            try:
                # Пытаемся распарсить как число
                num_val = float(val_str.replace(",", "."))
                if math.isnan(num_val):
                    return (1, "") if self.sort_reverse else (0, "")
                return (0, num_val)
            except:
                # Текстовая сортировка
                return (0, val_str)
        
        try:
            sorted_records = sorted(records, key=get_sort_key, reverse=self.sort_reverse)
            return sorted_records
        except Exception as e:
            print(f"[WARNING] Ошибка сортировки: {e}")
            return records
    
    def _apply_filters(self, event=None):
        """Применение фильтров"""
        if self._all_records:
            self._load_apply_json(silent=True)
    
    def _reset_filters(self):
        """Сброс фильтров"""
        self.line_filter.set("Все")
        self.day_filter.set("Все")
        if hasattr(self, 'search_entry'):
            self.search_entry.delete(0, tk.END)
        self._apply_filters()
    
    def _sort_by_column(self, column_index: int, column_name: str):
        """Сортировка по колонке"""
        # Переключаем направление сортировки, если кликнули по той же колонке
        if self.sort_column == column_index:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column_index
            self.sort_reverse = False
        
        # Обновляем индикаторы сортировки в заголовках
        for i, h in enumerate(HEADERS):
            heading_text = h
            if i == column_index:
                arrow = " ▼" if self.sort_reverse else " ▲"
                heading_text = h + arrow
            self.tree.heading(f"c{i}", text=heading_text)
        
        # Сортируем данные
        self._load_apply_json(silent=True)
    
    def _export_data(self):
        """Экспорт данных в Excel/CSV"""
        if not self._rows:
            messagebox.showinfo("Информация", "Нет данных для экспорта")
            return
        
        try:
            import csv
            from datetime import datetime
            
            # Запрашиваем путь для сохранения
            filename = filedialog.asksaveasfilename(
                title="Сохранить данные",
                defaultextension=".csv",
                filetypes=[
                    ("CSV файлы", "*.csv"),
                    ("Excel файлы", "*.xlsx"),
                    ("Все файлы", "*.*")
                ]
            )
            
            if not filename:
                return
            
            if filename.endswith('.xlsx'):
                # Экспорт в Excel (требует openpyxl)
                try:
                    from openpyxl import Workbook
                    wb = Workbook()
                    ws = wb.active
                    
                    # Заголовки
                    ws.append(HEADERS)
                    
                    # Данные
                    for row in self._rows:
                        ws.append(row)
                    
                    wb.save(filename)
                    messagebox.showinfo("Успех", f"Данные экспортированы в {filename}")
                except ImportError:
                    messagebox.showerror("Ошибка", 
                        "Для экспорта в Excel требуется библиотека openpyxl.\n"
                        "Установите её командой: pip install openpyxl")
            else:
                # Экспорт в CSV
                with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f, delimiter=';')
                    writer.writerow(HEADERS)
                    writer.writerows(self._rows)
                messagebox.showinfo("Успех", f"Данные экспортированы в {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать данные:\n{e}")
    
    # ---------- Отображение простоев ----------
    def _add_downtimes_to_tree(self, parent_id, downtimes: List[Dict]):
        """Добавление деталей простоев в дерево"""
        for idx, dt_item in enumerate(downtimes):
            if isinstance(dt_item, dict):
                # Извлекаем данные о простоях
                category = dt_item.get("category", dt_item.get("type", ""))
                reason = dt_item.get("reason", dt_item.get("cause", ""))
                description = dt_item.get("description", dt_item.get("comment", dt_item.get("details", "")))
                start = dt_item.get("start", dt_item.get("start_time", ""))
                end = dt_item.get("end", dt_item.get("end_time", ""))
                
                # Получаем длительность из данных или вычисляем по времени
                duration = _num(dt_item.get("duration_min", dt_item.get("duration", 0)))
                
                # Если длительность не указана или равна 0, вычисляем по времени начала и окончания
                if (duration == 0 or math.isnan(duration)) and start and end:
                    duration = _minutes_from_hhmm(start, end)
                
                # Формируем текст для отображения с категорией и описанием
                category_text = category if category else "—"
                reason_text = reason if reason else "Причина не указана"
                description_text = description if description else ""
                
                # Формируем строку с деталями простоя
                # Используем колонки для красивого отображения категории и описания
                # Комбинируем категорию и причину в одну строку для лучшей читаемости
                category_reason = f"{category_text}" if category_text and category_text != "—" else ""
                if reason_text and reason_text != "Причина не указана":
                    if category_reason:
                        category_reason += f" | {reason_text}"
                    else:
                        category_reason = reason_text
                
                # Если нет категории и причины, используем описание или название по умолчанию
                if not category_reason:
                    category_reason = description_text if description_text else "Простой"
                
                # Для отображения в колонке "Продукт" показываем комбинацию категория + причина
                product_display = category_reason if category_reason else "Простой"
                
                # В описание выносим дополнительную информацию (если описание не совпадает с причиной)
                if description_text and description_text != reason_text:
                    description_display = description_text
                else:
                    description_display = ""
                
                downtime_row = [
                    "",  # Job ID
                    product_display,  # Продукт - категория | причина
                    "",  # Линия - оставляем пустым для визуального отступа
                    "",  # День
                    "",  # Смена
                    start or "",  # Начало
                    end or "",  # Конец
                    _fmt(duration),  # Длит (мин) - длительность простоя
                    "",  # Σ простой (мин)
                    "",  # % простоя
                    "",  # Событий
                    "",  # План. простой (мин)
                    "",  # EffMin (мин)
                    "",  # Ном. скорость (ш)
                    "",  # Потолок (шт)
                    description_display,  # Факт (шт) - описание простоя
                    ""   # OEE, %
                ]
                
                # Формируем всплывающую подсказку с полной информацией
                tooltip_text = f"Категория: {category_text}\n"
                tooltip_text += f"Причина: {reason_text}\n"
                if description_text:
                    tooltip_text += f"Описание: {description_text}\n"
                tooltip_text += f"Начало: {start or '—'}\n"
                tooltip_text += f"Конец: {end or '—'}\n"
                tooltip_text += f"Длительность: {_fmt(duration)} мин"
                
                # Добавляем визуальный отступ и иконку для простоев
                item_id = self.tree.insert(parent_id, "end", text="  └─", values=downtime_row,
                               tags=("downtime_detail",))
                
                # Сохраняем текст подсказки в словаре для использования при наведении
                self._tooltips[item_id] = tooltip_text
    
    def _show_downtime_details(self, item):
        """Показать окно с деталями простоя"""
        tooltip_text = self._tooltips.get(item)
        if not tooltip_text:
            return
        
        # Закрываем предыдущее окно, если есть
        if self._tooltip_window:
            try:
                self._tooltip_window.destroy()
            except:
                pass
        
        # Создаем красивое окно с деталями
        win = tk.Toplevel(self._tab)
        win.title("Детали простоя")
        win.transient(self._tab.winfo_toplevel())
        win.grab_set()
        
        # Позиционирование окна
        try:
            x = self._tab.winfo_rootx() + 100
            y = self._tab.winfo_rooty() + 100
            win.geometry(f"500x350+{x}+{y}")
        except:
            win.geometry("500x350")
        
        win.resizable(True, True)
        win.minsize(400, 250)
        
        # Основной фрейм
        main_frame = ttk.Frame(win, padding=20)
        main_frame.pack(fill="both", expand=True)
        
        # Заголовок
        header_label = ttk.Label(main_frame, text="Детали простоя", 
                               font=("TkDefaultFont", 12, "bold"))
        header_label.pack(anchor="w", pady=(0, 15))
        
        # Фрейм для деталей
        details_frame = ttk.LabelFrame(main_frame, text="Информация", padding=15)
        details_frame.pack(fill="both", expand=True)
        
        # Парсим текст подсказки для красивого отображения
        lines = tooltip_text.split("\n")
        details = {}
        for line in lines:
            if ":" in line:
                key, value = line.split(":", 1)
                details[key.strip()] = value.strip()
        
        # Отображаем детали в виде формы
        row = 0
        labels_config = [
            ("Категория", details.get("Категория", "—")),
            ("Причина", details.get("Причина", "—")),
            ("Начало", details.get("Начало", "—")),
            ("Конец", details.get("Конец", "—")),
            ("Длительность", details.get("Длительность", "—")),
        ]
        
        for label_text, value_text in labels_config:
            ttk.Label(details_frame, text=f"{label_text}:", font=("TkDefaultFont", 9, "bold")).grid(
                row=row, column=0, sticky="ne", padx=(0, 10), pady=5)
            ttk.Label(details_frame, text=value_text, font=("TkDefaultFont", 9)).grid(
                row=row, column=1, sticky="w", pady=5)
            row += 1
        
        # Описание отдельно, если есть
        description = details.get("Описание", "")
        if description and description != "—":
            ttk.Label(details_frame, text="Описание:", font=("TkDefaultFont", 9, "bold")).grid(
                row=row, column=0, sticky="ne", padx=(0, 10), pady=(10, 5))
            
            # Текстовое поле для описания с переносами
            desc_frame = ttk.Frame(details_frame)
            desc_frame.grid(row=row, column=1, sticky="nsew", pady=(10, 5))
            
            desc_text = tk.Text(desc_frame, height=4, wrap="word", 
                              font=("TkDefaultFont", 9), relief="flat",
                              background="#f5f5f5", borderwidth=1)
            desc_text.insert("1.0", description)
            desc_text.config(state="disabled")
            desc_text.pack(fill="both", expand=True)
            
            row += 1
        
        details_frame.grid_columnconfigure(1, weight=1)
        
        # Кнопка закрытия
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=(15, 0))
        
        def close_window():
            win.destroy()
            self._tooltip_window = None
        
        ttk.Button(btn_frame, text="Закрыть", command=close_window).pack(side="right")
        
        win.protocol("WM_DELETE_WINDOW", close_window)
        
        self._tooltip_window = win
        self._tooltip_item = item
    
    def _on_mouse_motion(self, event):
        """Обработка движения мыши - убрано автоматическое показывание tooltip"""
        # Убрано автоматическое показывание tooltip - теперь показывается по клику
        pass
    
    def _on_mouse_leave(self, event):
        """Обработка выхода мыши из таблицы"""
        # Не закрываем окно при выходе мыши - только по кнопке
        pass
    
    def _on_row_click(self, event):
        """Обработка клика для показа деталей простоя"""
        try:
            # Небольшая задержка, чтобы не конфликтовать с выделением строки
            item = self.tree.identify_row(event.y)
            if item:
                self._tab.after(200, lambda i=item: self._check_and_show_downtime(i))
        except:
            pass
    
    def _check_and_show_downtime(self, item):
        """Проверка и показ деталей простоя"""
        try:
            # Проверяем, является ли это дочерним элементом (простоем)
            parent = self.tree.parent(item)
            if parent and item in self._tooltips:
                # Это строка простоя - показываем окно с деталями
                self._show_downtime_details(item)
        except:
            pass
    
    def _on_row_double_click(self, event):
        """Обработка двойного клика для раскрытия/сворачивания простоев"""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if not item:
            return
        
        # Если это дочерний элемент (простой), не обрабатываем двойной клик
        parent = self.tree.parent(item)
        if parent:
            return
        
        # Проверяем, есть ли простои для этой строки
        values = self.tree.item(item, "values")
        if not values or len(values) < 1:
            return
        
        job_id = values[0] if values else ""
        line = values[2] if len(values) > 2 else ""
        day = values[3] if len(values) > 3 else ""
        
        row_id = f"{job_id}_{line}_{day}"
        downtimes = self.downtimes_data.get(row_id, [])
        
        if not downtimes:
            return
        
        # Переключаем раскрытие
        if item in self.expanded_rows:
            # Сворачиваем - удаляем дочерние элементы
            for child in list(self.tree.get_children(item)):
                self.tree.delete(child)
            self.tree.item(item, text="▶")
            self.expanded_rows.discard(item)
        else:
            # Раскрываем - добавляем простои
            self._add_downtimes_to_tree(item, downtimes)
            self.tree.item(item, text="▼")
            self.expanded_rows.add(item)
    
    def _toggle_downtimes(self):
        """Переключение отображения простоев"""
        if self.show_downtimes_var.get():
            # Показать все простои
            for item in self.tree.get_children():
                values = self.tree.item(item, "values")
                if not values or len(values) < 1:
                    continue
                
                job_id = values[0] if values else ""
                line = values[2] if len(values) > 2 else ""
                day = values[3] if len(values) > 3 else ""
                
                row_id = f"{job_id}_{line}_{day}"
                downtimes = self.downtimes_data.get(row_id, [])
                
                if downtimes and item not in self.expanded_rows:
                    self._add_downtimes_to_tree(item, downtimes)
                    self.tree.item(item, text="▼")
                    self.expanded_rows.add(item)
        else:
            # Скрыть все простои
            for item in list(self.expanded_rows):
                for child in list(self.tree.get_children(item)):
                    self.tree.delete(child)
                self.tree.item(item, text="▶")
            self.expanded_rows.clear()
    
    # ---------- Статистика ----------
    def _update_statistics(self, records: List[Dict]):
        """Обновление статистики"""
        if not records:
            # Обновляем карточки нулями
            if hasattr(self, 'card_records'):
                self.card_records.value_label.config(text="0")
            if hasattr(self, 'card_oee'):
                self.card_oee.value_label.config(text="— %", foreground="#666")
            if hasattr(self, 'card_downtimes'):
                self.card_downtimes.value_label.config(text="0")
            if hasattr(self, 'card_downtime_min'):
                self.card_downtime_min.value_label.config(text="0 мин")
            if hasattr(self, 'lbl_status'):
                self.lbl_status.config(text="● Нет данных", foreground="#666")
            return
        
        total_records = len(records)
        total_downtimes = 0
        total_downtime_min = 0.0
        avg_oee = 0.0
        oee_count = 0
        
        for r in records:
            downtimes = r.get("downtimes", [])
            if isinstance(downtimes, list):
                total_downtimes += len(downtimes)
                for dt in downtimes:
                    if isinstance(dt, dict):
                        total_downtime_min += _num(dt.get("duration_min", dt.get("duration", 0)))
            
            # Вычисляем OEE
            dur = _num(r.get("duration_min", 0))
            speed = _num(r.get("speed", 0))
            fact = _num(r.get("fact", 0))
            sum_dt = 0.0
            plan_dt = _num(r.get("planned_downtime_min", 0))
            
            for dt_item in r.get("downtimes", []):
                if isinstance(dt_item, dict):
                    sum_dt += _num(dt_item.get("duration_min", dt_item.get("duration", 0)))
            
            effmin = max(0, (dur or 0) - sum_dt - plan_dt)
            ceil_units = effmin * speed / 60 if speed and effmin else 0
            oee = (fact / ceil_units * 100) if (ceil_units and ceil_units > 0) else 0
            
            if not math.isnan(oee) and oee > 0:
                avg_oee += oee
                oee_count += 1
        
        avg_oee = avg_oee / oee_count if oee_count > 0 else 0
        
        # Обновляем карточки статистики
        if hasattr(self, 'card_records'):
            self.card_records.value_label.config(text=str(total_records))
        
        if hasattr(self, 'card_oee'):
            if avg_oee > 0:
                oee_text = f"{_fmt(avg_oee, 1)}%"
                oee_color = "#28a745" if avg_oee >= 85 else "#ffc107" if avg_oee >= 70 else "#dc3545"
                self.card_oee.value_label.config(text=oee_text, foreground=oee_color)
            else:
                self.card_oee.value_label.config(text="— %", foreground="#666")
        
        if hasattr(self, 'card_downtimes'):
            self.card_downtimes.value_label.config(text=str(total_downtimes))
        
        if hasattr(self, 'card_downtime_min'):
            self.card_downtime_min.value_label.config(text=f"{_fmt(total_downtime_min)} мин")
        
        # Обновляем статус в шапке
        if hasattr(self, 'lbl_status'):
            if total_records > 0:
                self.lbl_status.config(text="● Загружено", foreground="#28a745")
            else:
                self.lbl_status.config(text="● Готов", foreground="#666")

# ===== точка входа =====
# было: def show_json_import_tab(nb: ttk.Notebook):
def show_json_import_tab(nb: ttk.Notebook, on_import=None):
    JsonImportTab(nb, on_import=on_import)

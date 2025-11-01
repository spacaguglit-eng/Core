# -*- coding: utf-8 -*-
from __future__ import annotations
from tkinter import ttk
import tkinter as tk
import tkinter.font as tkFont
import tkinter.messagebox as MB
import json, os, re, uuid
from typing import Dict, List, Tuple
from product_parse import parse_product_name, clear_product_parse_cache

_THIS_DIR = os.path.dirname(__file__)
_RULES_JSON_LEGACY     = os.path.join(_THIS_DIR, "rules_data.json")
_EVICTIONS_JSON_LEGACY = os.path.join(_THIS_DIR, "evictions_data.json")
_RULES_SETS_JSON       = os.path.join(_THIS_DIR, "rules_sets.json")
_EVICT_SETS_JSON       = os.path.join(_THIS_DIR, "evictions_sets.json")
_NORMS_JSON            = os.path.join(_THIS_DIR, "norms_data.json")

def _low(s: str) -> str:
    return (s or "").replace("\xa0", " ").strip().lower().replace("ё", "е")

def _load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def _save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _vol_to_ml(s: str) -> int | None:
    if not s:
        return None
    x = _low(s).replace(" ", "").replace(",", ".")
    m = re.search(r"(\d+(?:\.\d+)?)", x)
    if not m:
        return None
    num = float(m.group(1))
    if "мл" in x or "ml" in x:
        return int(round(num))
    if "л" in x or re.search(r"(?<!m)l\b", x):
        return int(round(num * 1000))
    return int(round(num))

def _wrap_to_pixels(text: str, font: tkFont.Font, max_px: int) -> str:
    if not text: return ""
    words = str(text).split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if not cur or font.measure(trial) <= max_px:
            cur = trial
        else:
            lines.append(cur); cur = w
    if cur: lines.append(cur)
    return "\n".join(lines)

def _unwrap(s: str) -> str:
    return (s or "").replace("\r", "").replace("\n", " ").strip()

def _make_grid_tab(parent, columns, headers, widths):
    tab = ttk.Frame(parent)
    bar = ttk.Frame(tab); bar.pack(fill="x", padx=8, pady=(8,4))
    btn_add  = ttk.Button(bar, text="Добавить");  btn_add.pack(side="left")
    btn_edit = ttk.Button(bar, text="Изменить");  btn_edit.pack(side="left", padx=6)
    btn_del  = ttk.Button(bar, text="Удалить");   btn_del.pack(side="left", padx=6)
    btn_save = ttk.Button(bar, text="Сохранить"); btn_save.pack(side="left", padx=12)
    btn_load = ttk.Button(bar, text="Загрузить"); btn_load.pack(side="left", padx=6)
    info = ttk.Label(bar, text="", foreground="#666"); info.pack(side="left", padx=12)
    wrap = ttk.Frame(tab); wrap.pack(fill="both", expand=True, padx=8, pady=(0,8))
    tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="extended",
                        style="Wrap.Treeview")
    vs = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vs.set)
    for cid, hdr in zip(columns, headers):
        tree.heading(cid, text=hdr)
    for cid, w in zip(columns, widths):
        tree.column(cid, width=w, anchor=("w" if cid in ("from","to","exceptions") else "center"))
    tree.grid(row=0, column=0, sticky="nsew")
    vs.grid(row=0, column=1, sticky="ns")
    wrap.rowconfigure(0, weight=1); wrap.columnconfigure(0, weight=1)
    style = ttk.Style(tree)
    style.configure("Wrap.Treeview", rowheight=56)
    _f_cell = tkFont.Font(family="Segoe UI", size=9)
    _wrap_px = {cid: max(40, tree.column(cid, option="width") - 12) for cid in columns}
    _raw_by_iid: dict[str, dict[str, str]] = {}
    def _display_values(raw_vals: dict[str, str]) -> tuple[str, ...]:
        wrap_cols = {"from", "to", "exceptions", "product", "CIP1", "CIP2", "CIP3"}
        out: list[str] = []
        for cid in columns:
            txt = raw_vals.get(cid, "")
            if cid in wrap_cols:
                out.append(_wrap_to_pixels(txt, _f_cell, _wrap_px.get(cid, 160)))
            else:
                out.append(str(txt))
        return tuple(out)
    def _insert_row(values: dict[str, str]) -> str:
        raw_vals = {c: _unwrap(values.get(c, "")) for c in columns}
        disp = _display_values(raw_vals)
        iid = tree.insert("", "end", values=disp)
        _raw_by_iid[iid] = raw_vals
        return iid
    def _update_row(iid: str, values: dict[str, str]) -> None:
        raw_vals = {c: _unwrap(values.get(c, _raw_by_iid.get(iid, {}).get(c, ""))) for c in columns}
        _raw_by_iid[iid] = raw_vals
        disp = _display_values(raw_vals)
        for j, cid in enumerate(columns):
            tree.set(iid, cid, disp[j])
    def _edit_dialog(iid: str | None = None):
        is_new = iid is None
        cur = {c: "" for c in columns}
        if iid is not None:
            cur = {c: _raw_by_iid.get(iid, {}).get(c, "") for c in columns}
        win = tk.Toplevel(tab); win.title("Редактирование")
        win.transient(tab.winfo_toplevel()); win.grab_set()
        # Удобные размеры и возможность растягивать
        try:
            px = tab.winfo_rootx() + 40; py = tab.winfo_rooty() + 40
            win.geometry(f"780x520+{px}+{py}")
        except Exception:
            win.geometry("780x520")
        win.minsize(640, 420)
        win.resizable(True, True)
        frm = ttk.Frame(win); frm.pack(fill="both", expand=True, padx=12, pady=12)
        widgets = {}
        # Столбцы: метка слева, поле справа
        for i, (cid, hdr) in enumerate(zip(columns, headers)):
            ttk.Label(frm, text=hdr + ":").grid(row=i, column=0, sticky="ne", padx=(0,8), pady=4)
            # Длинные поля редактируем в многострочном Text со скроллбаром
            is_long = cid in ("exceptions","from","to","product","CIP1","CIP2","CIP3")
            if is_long:
                wrap_frame = ttk.Frame(frm); wrap_frame.grid(row=i, column=1, sticky="nsew", pady=4)
                txt = tk.Text(wrap_frame, height=6, wrap="word")
                txt.insert("1.0", cur.get(cid, ""))
                vsb = ttk.Scrollbar(wrap_frame, orient="vertical", command=txt.yview)
                txt.configure(yscrollcommand=vsb.set)
                txt.pack(side="left", fill="both", expand=True)
                vsb.pack(side="right", fill="y")
                widgets[cid] = txt
            else:
                ent = ttk.Entry(frm)
                ent.insert(0, cur.get(cid, ""))
                ent.grid(row=i, column=1, sticky="ew", pady=4)
                widgets[cid] = ent
            frm.grid_columnconfigure(1, weight=1)
        btns = ttk.Frame(frm); btns.grid(row=len(columns)+1, column=0, columnspan=2, sticky="e", pady=(8,0))
        def _collect():
            vals = {}
            for cid, w in widgets.items():
                vals[cid] = w.get("1.0", "end-1c") if isinstance(w, tk.Text) else w.get()
            return vals
        def _ok():
            vals = _collect()
            if is_new: _insert_row(vals)
            else:      _update_row(iid, vals)
            win.destroy()
        ttk.Button(btns, text="OK", command=_ok).pack(side="right", padx=6)
        ttk.Button(btns, text="Отмена", command=win.destroy).pack(side="right")
        win.bind("<Return>", lambda _e: _ok())
        win.bind("<Escape>", lambda _e: win.destroy())
    def _add():
        _edit_dialog(None)
    def _edit_selected():
        sel = tree.selection()
        if not sel:
            return
        _edit_dialog(sel[0])
    def _del():
        for iid in tree.selection():
            _raw_by_iid.pop(iid, None)
            tree.delete(iid)
    btn_add.configure(command=_add)
    btn_edit.configure(command=_edit_selected)
    btn_del.configure(command=_del)
    def _on_dbl_click(event):
        iid = tree.identify_row(event.y)
        if iid:
            _edit_dialog(iid)
    tree.bind("<Double-1>", _on_dbl_click)
    def _save_json_rows(path: str):
        rows = []
        for iid in tree.get_children(""):
            raw = _raw_by_iid.get(iid, {})
            rows.append({c: raw.get(c, "") for c in columns})
        _save_json(path, rows)
    def _load_json_rows(path: str):
        for iid in tree.get_children(""): tree.delete(iid)
        _raw_by_iid.clear()
        rows = _load_json(path, [])
        for r in rows:
            _insert_row({c: r.get(c, "") for c in columns})
        info.config(text=f"Загружено: {len(rows)}")
    tab._save_json = _save_json_rows
    tab._load_json = _load_json_rows
    return tab, tree, btn_save, btn_load, info

def _cell_mark(a: Dict[str, str], b: Dict[str, str]) -> str:
    def _canon_brand(s: str) -> str:
        x = _low(s); x = re.sub(r'\b(тм|tm|®|™)\b', '', x)
        x = re.sub(r'["«»\'`]+', ' ', x); x = re.sub(r'\s*\([^)]*\)\s*', ' ', x)
        x = re.sub(r'\s+', ' ', x); return x.strip()
    def _canon_flavor(s: str) -> str:
        x = _low(s); x = re.sub(r'\s+', ' ', x).strip(); return x
    af, bf = _canon_flavor(a.get("flavor", "")), _canon_flavor(b.get("flavor", ""))
    ab, bb = _canon_brand(a.get("brand", "")), _canon_brand(b.get("brand", ""))
    av_ml = _vol_to_ml(a.get("volume")) or _vol_to_ml(a.get("container"))
    bv_ml = _vol_to_ml(b.get("volume")) or _vol_to_ml(b.get("container"))
    same_v = (av_ml is not None and bv_ml is not None and av_ml == bv_ml)
    diff_v = (av_ml is not None and bv_ml is not None and av_ml != bv_ml)
    if diff_v: return "П"
    if same_v and af and (af == bf) and (ab != bb): return "С"
    return ""

def _collect_catalog_flavors(rows_cache: List[dict]) -> set[str]:
    clear_product_parse_cache()
    flavors: set[str] = set()
    for r in rows_cache:
        name = (r.get("name") or "").strip()
        cont = (r.get("container") or "").strip()
        if not name: continue
        meta = parse_product_name(name, cont)
        f = _low(meta.get("flavor") or "")
        if f: flavors.add(re.sub(r"\s+", " ", f).strip())
    return flavors

def _show_mismatch_report(title: str, mismatches: list[dict]):
    win = tk.Toplevel(); win.title(title); win.geometry("860x520")
    ttk.Label(win, text=f"Найдено несоответствий: {len(mismatches)}", foreground="#666")\
        .pack(anchor="w", padx=8, pady=(8,4))
    cols = ("row","column","value","issue")
    headers = {"row":"Строка","column":"Колонка","value":"Значение","issue":"Проблема"}
    widths  = {"row":70,"column":120,"value":360,"issue":260}
    tree = ttk.Treeview(win, columns=cols, show="headings", height=20); tree.pack(fill="both", expand=True, padx=8, pady=8)
    for c in cols:
        tree.heading(c, text=headers[c]); tree.column(c, width=widths[c], anchor=("center" if c in ("row","column") else "w"))
    for m in mismatches:
        tree.insert("", "end", values=(m["row"], m["column"], m["value"], m["issue"]))
    def _copy_csv():
        import csv, io
        buf = io.StringIO(); w = csv.writer(buf, delimiter=";")
        w.writerow([headers[c] for c in cols])
        for m in mismatches: w.writerow([m["row"], m["column"], m["value"], m["issue"]])
        win.clipboard_clear(); win.clipboard_append(buf.getvalue()); win.update()
    ttk.Button(win, text="Скопировать как CSV (;)", command=_copy_csv)\
        .pack(anchor="e", padx=8, pady=(0,8))

def _ensure_rules_sets() -> list[dict]:
    if not os.path.isfile(_RULES_SETS_JSON):
        legacy = _load_json(_RULES_JSON_LEGACY, [])
        sets = [{"id":"default","name":"Default","lines":[],"rules":legacy}]
        _save_json(_RULES_SETS_JSON, sets)
    return _load_json(_RULES_SETS_JSON, [])

def _ensure_evict_sets() -> list[dict]:
    if not os.path.isfile(_EVICT_SETS_JSON):
        legacy = _load_json(_EVICTIONS_JSON_LEGACY, [])
        sets = [{"id":"default","name":"Default","lines":[],"rules":legacy}]
        _save_json(_EVICT_SETS_JSON, sets)
    return _load_json(_EVICT_SETS_JSON, [])

def _active_set_for_line(sets: list[dict], line: str) -> tuple[dict|None, str|None]:
    ln = _low(line)
    owners = [s for s in sets if any(_low(x) == ln for x in (s.get("lines") or []))]
    if not owners:
        return None, f"Для линии «{line}» не назначен набор правил."
    if len(owners) > 1:
        names = ", ".join(s.get("name","?") for s in owners)
        return None, f"Линия «{line}» привязана к нескольким наборам: {names}."
    return owners[0], None

def _save_rules_sets(sets: list[dict]):
    _save_json(_RULES_SETS_JSON, sets)

def _save_evict_sets(sets: list[dict]):
    _save_json(_EVICT_SETS_JSON, sets)

def _lines_in_use(sets: list[dict]) -> dict[str, str]:
    m = {}
    for s in sets:
        for ln in (s.get("lines") or []):
            m[ln] = s["id"]
    return m

def _prod_key_from_rule(text: str) -> str:
    meta = parse_product_name(text or "", "")
    t = _low(meta.get("type") or "")
    f = _low(meta.get("flavor") or "")
    if not f:
        return ""
    return f"{t}::{f}" if t else f"::{f}"

def _prod_key_from_catalog(name: str, container: str = "") -> str:
    meta = parse_product_name(name or "", container or "")
    t = _low(meta.get("type") or "")
    f = _low(meta.get("flavor") or "")
    if not f:
        return ""
    return f"{t}::{f}" if t else f"::{f}"

def _collect_catalog_prodkeys(rows_cache: List[dict]) -> set[str]:
    clear_product_parse_cache()
    keys: set[str] = set()
    for r in rows_cache:
        name = (r.get("name") or "").strip()
        cont = (r.get("container") or "").strip()
        if not name:
            continue
        k = _prod_key_from_catalog(name, cont)
        if k:
            keys.add(k)
    return keys

def _parse_sip_set_row(row: dict) -> tuple[str, dict] | None:
    if not row:
        return None
    prod_key = _prod_key_from_rule(row.get("product", ""))
    if not prod_key:
        return None
    vals = {k: (row.get(k) or "").strip() for k in ("CIP1", "CIP2", "CIP3")}
    base = None
    for k, v in vals.items():
        if _low(v) in ("база", "base"):
            if base is not None:
                return None
            base = k
    if base is None:
        return None
    alts = {"CIP1": set(), "CIP2": set(), "CIP3": set()}
    for cip_name in ("CIP1", "CIP2", "CIP3"):
        if cip_name == base:
            continue
        raw = vals.get(cip_name, "")
        if not raw:
            continue
        tokens = [t.strip() for t in raw.split(";") if t.strip()]
        for t in tokens:
            k = _prod_key_from_rule(t)
            if k:
                alts[cip_name].add(k)
    return prod_key, {"base": base, "alts": alts}

def _build_sip_map_for_line(line: str) -> dict[str, dict] | None:
    sets = _ensure_rules_sets()
    s, err = _active_set_for_line(sets, line)
    if err or not s or not (s.get("rules") or []):
        return None
    out = {}
    for r in (s.get("rules") or []):
        parsed = _parse_sip_set_row(r)
        if not parsed:
            continue
        key, val = parsed
        out[key] = val
    return out

def _parse_evict_set_row(row: dict) -> tuple[str, set[str], set[str]] | None:
    if not row:
        return None
    fk = _prod_key_from_rule(row.get("from", ""))
    if not fk:
        return None

    # --- поддержка нескольких 'to' через ';' ---
    to_raw = (row.get("to") or "").strip()
    to_set: set[str] = set()
    if to_raw:
        for t in [t.strip() for t in to_raw.split(";") if t.strip()]:
            k = _prod_key_from_rule(t)
            if k:
                to_set.add(k)
    if not to_set:
        return None

    # исключения как и раньше (также через ';')
    exc_raw = (row.get("exceptions") or "").strip()
    exc: set[str] = set()
    if exc_raw:
        for t in [t.strip() for t in exc_raw.split(";") if t.strip()]:
            k = _prod_key_from_rule(t)
            if k:
                exc.add(k)

    return fk, to_set, exc


def _build_evict_maps_for_line(line: str) -> tuple[set[tuple[str,str]], set[tuple[str,str]]]:
    sets = _ensure_evict_sets()
    s, err = _active_set_for_line(sets, line)
    if err or not s or not (s.get("rules") or []):
        return set(), set()

    allow: set[tuple[str,str]] = set()
    deny:  set[tuple[str,str]] = set()

    for r in (s.get("rules") or []):
        parsed = _parse_evict_set_row(r)
        if not parsed:
            continue
        fk, to_set, exc = parsed
        # все цели «В» срабатывают
        for tk in to_set:
            allow.add((fk, tk))
        # исключения блокируют конкретные пары (from, exception)
        for exk in exc:
            deny.add((fk, exk))

    return allow, deny


def show_matrix_tab(nb, catalog):
    try:
        for _tid in list(nb.tabs()):
            if nb.tab(_tid, "text") in ("Матрицы", "Правила (наборы)", "Вытеснения (наборы)", "Нормативы"):
                nb.forget(_tid)
    except Exception:
        pass
    tab_matrix = ttk.Frame(nb); nb.add(tab_matrix, text="Матрицы")
    rows_cache: List[dict] = list(catalog.rows())
    lines = sorted({(r.get("line", "") or "").strip() for r in rows_cache if r.get("line")})
    frm_top = ttk.Frame(tab_matrix); frm_top.pack(fill="x", padx=8, pady=(8,4))
    SELECTED_LINES = set(lines)
    btn_lines = ttk.Button(frm_top, text="Линии ▾"); btn_lines.pack(side="left")
    lbl_sel = ttk.Label(frm_top, text="(все)"); lbl_sel.pack(side="left", padx=8)
    def _open_lines_panel():
        if not lines: return
        win = tk.Toplevel(tab_matrix); win.title("Выбор линий")
        win.transient(tab_matrix); win.resizable(False, True); win.attributes("-topmost", True)
        lb = tk.Listbox(win, selectmode="extended", height=min(12, max(6,len(lines))))
        for ln in lines: lb.insert("end", ln)
        for i,ln in enumerate(lines):
            if ln in SELECTED_LINES: lb.selection_set(i)
        lb.pack(fill="both", expand=True, padx=8, pady=8)
        def _ok():
            chosen = {lines[i] for i in lb.curselection()}
            SELECTED_LINES.clear(); SELECTED_LINES.update(chosen or [])
            lbl_sel.config(text=(",".join(sorted(SELECTED_LINES)) if SELECTED_LINES else "—"))
            build_all_matrices()
            win.destroy()
        ttk.Button(win, text="OK", command=_ok).pack(anchor="e", padx=8, pady=(0,8))
        win.grab_set()
    btn_lines.configure(command=_open_lines_panel)
    def _sip_key(meta: dict) -> str:
        return f"{_low(meta.get('type',''))}::{_low(meta.get('flavor',''))}".strip(":")
    def _debug_show_line_view():
        chosen = sorted(SELECTED_LINES) if SELECTED_LINES else []
        if not chosen:
            MB.showinfo("Проверка линии", "Сначала выберите линии.")
            return
        win = tk.Toplevel(tab_matrix)
        win.title("Проверка ключей CIP (что видит построитель)")
        win.geometry("1100x520")
        win.transient(tab_matrix)
        info_lbl = ttk.Label(win, text="", foreground="#666")
        info_lbl.pack(anchor="w", padx=8, pady=(8, 4))
        cols = ("line", "name", "container", "type", "flavor", "key", "in_rules", "base", "excs")
        headers = {
            "line": "Линия",
            "name": "Наименование",
            "container": "Тара/объём",
            "type": "Тип",
            "flavor": "Вкус",
            "key": "Ключ (тип::вкус)",
            "in_rules": "Есть правило?",
            "base": "База",
            "excs": "Искл. (шт)",
        }
        tree = ttk.Treeview(win, columns=cols, show="headings", height=20)
        vs = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(0, 8))
        vs.pack(side="left", fill="y", padx=(0, 8), pady=(0, 8))
        tree.column("line", width=100, anchor="center")
        tree.column("name", width=320, anchor="w")
        tree.column("container", width=120, anchor="center")
        tree.column("type", width=90, anchor="center")
        tree.column("flavor", width=160, anchor="w")
        tree.column("key", width=160, anchor="w")
        tree.column("in_rules", width=110, anchor="center")
        tree.column("base", width=80, anchor="center")
        tree.column("excs", width=100, anchor="center")
        for c in cols:
            tree.heading(c, text=headers[c])
        total, matched = 0, 0
        for line in chosen:
            sip_map = _build_sip_map_for_line(line) or {}
            nline = _low(line)
            rows = [r for r in rows_cache if _low(r.get("line","")) == nline and r.get("name")]
            for r in rows:
                total += 1
                name = (r.get("name") or "").strip()
                cont = (r.get("container") or "").strip()
                meta = parse_product_name(name, cont)
                key = _sip_key(meta)
                hit = key in sip_map
                if hit: matched += 1
                base = sip_map.get(key, {}).get("base", "")
                alts_map = sip_map.get(key, {}).get("alts", {}) if hit else {}
                ex_cnt = sum(len(v) for v in (alts_map or {}).values())
                tree.insert(
                    "", "end",
                    values=(line, name, cont, meta.get("type",""), meta.get("flavor",""),
                            key, ("Да" if hit else "Нет"), base, ex_cnt),
                    tags=("hit",) if hit else ("miss",)
                )
        tree.tag_configure("hit", background="#eaffea")
        tree.tag_configure("miss", background="#ffecec")
        info_lbl.config(text=f"Строк каталога: {total} • Совпало с правилами: {matched} • Линий: {len(chosen)}")
        def _copy_csv():
            import csv, io
            buf = io.StringIO()
            w = csv.writer(buf, delimiter=";")
            w.writerow([headers[c] for c in cols])
            for iid in tree.get_children(""):
                w.writerow(tree.item(iid, "values"))
            win.clipboard_clear(); win.clipboard_append(buf.getvalue()); win.update()
        ttk.Button(win, text="Скопировать как CSV (;)", command=_copy_csv)\
            .pack(anchor="e", padx=8, pady=(0, 8))
    ttk.Button(frm_top, text="Проверить линию", command=_debug_show_line_view)\
        .pack(side="left", padx=8)
    
    # Панель навигации и управления
    nav_frame = ttk.Frame(tab_matrix)
    nav_frame.pack(fill="x", padx=8, pady=(8, 4))
    
    # Кнопки навигации
    nav_buttons = ttk.Frame(nav_frame)
    nav_buttons.pack(side="left")
    
    def _nav_home():
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)
    
    def _nav_end():
        canvas.xview_moveto(1)
        canvas.yview_moveto(1)
    
    def _nav_left():
        canvas.xview_scroll(-1, "units")
    
    def _nav_right():
        canvas.xview_scroll(1, "units")
    
    def _nav_up():
        canvas.yview_scroll(-1, "units")
    
    def _nav_down():
        canvas.yview_scroll(1, "units")
    
    ttk.Button(nav_buttons, text="🏠", command=_nav_home, width=3).pack(side="left", padx=1)
    ttk.Button(nav_buttons, text="🔚", command=_nav_end, width=3).pack(side="left", padx=1)
    ttk.Button(nav_buttons, text="⬅", command=_nav_left, width=3).pack(side="left", padx=1)
    ttk.Button(nav_buttons, text="➡", command=_nav_right, width=3).pack(side="left", padx=1)
    ttk.Button(nav_buttons, text="⬆", command=_nav_up, width=3).pack(side="left", padx=1)
    ttk.Button(nav_buttons, text="⬇", command=_nav_down, width=3).pack(side="left", padx=1)
    
    # Кнопки масштабирования
    zoom_buttons = ttk.Frame(nav_frame)
    zoom_buttons.pack(side="left", padx=(20, 0))
    
    ttk.Button(zoom_buttons, text="🔍+", command=lambda: _apply_zoom(SCALE["k"] * 1.1), width=4).pack(side="left", padx=1)
    ttk.Button(zoom_buttons, text="🔍-", command=lambda: _apply_zoom(SCALE["k"] / 1.1), width=4).pack(side="left", padx=1)
    ttk.Button(zoom_buttons, text="🔍1", command=lambda: _apply_zoom(1.0), width=4).pack(side="left", padx=1)
    
    # Кнопка центрирования
    ttk.Button(nav_frame, text="🎯 Центр", command=lambda: (canvas.xview_moveto(0.5), canvas.yview_moveto(0.5))).pack(side="left", padx=(20, 0))
    
    # Кнопка обновления матриц
    def _refresh_matrices():
        try:
            build_all_matrices()
            # Обновляем метку масштаба
            try:
                scale_label.config(text=f"Масштаб: {int(SCALE['k'] * 100)}%")
            except:
                pass
        except Exception as e:
            print(f"Ошибка при обновлении матриц: {e}")
    
    ttk.Button(nav_frame, text="🔄 Обновить", command=_refresh_matrices).pack(side="left", padx=(20, 0))
    
    # Информация о масштабе
    scale_label = ttk.Label(nav_frame, text="Масштаб: 100%")
    scale_label.pack(side="right")
    
    # Информационная панель
    info_frame = ttk.Frame(tab_matrix)
    info_frame.pack(fill="x", padx=8, pady=(0, 4))
    
    info_text = "💡 Навигация: 🏠🔚 - начало/конец | ⬅➡⬆⬇ - движение | 🔍+/- - масштаб | 🎯 - центр | Space+ЛКМ - перетаскивание"
    info_label = ttk.Label(info_frame, text=info_text, font=("Arial", 8), foreground="gray")
    info_label.pack(side="left")
    
    canvas = tk.Canvas(tab_matrix, highlightthickness=0, bg="white")
    vsb = ttk.Scrollbar(tab_matrix, orient="vertical", command=canvas.yview)
    hsb = ttk.Scrollbar(tab_matrix, orient="horizontal", command=canvas.xview)
        # === ПРОКРУТКА И ЗУМ =======================================  ### NEW
    SCALE = {"k": 1.00}          # текущий зум
    SCALE_MIN, SCALE_MAX = 0.5, 2.5
    WHEEL_STEP = 60               # пикселей за один тик колеса

    # Перетаскивание полотна (курсор-рука)
    def _scan_mark(e):
        canvas.scan_mark(e.x, e.y)
        canvas.configure(cursor="fleur")
    
    def _scan_drag(e):
        canvas.scan_dragto(e.x, e.y, gain=1)
    
    def _scan_release(_e):
        canvas.configure(cursor="")

    # Средняя кнопка мыши для перетаскивания
    canvas.bind("<ButtonPress-2>", _scan_mark)
    canvas.bind("<B2-Motion>", _scan_drag)
    canvas.bind("<ButtonRelease-2>", _scan_release)
    
    # Space + ЛКМ для перетаскивания
    canvas.bind("<space>", lambda e: canvas.configure(cursor="fleur"))
    canvas.bind("<ButtonPress-1>", lambda e: (canvas.scan_mark(e.x, e.y) if canvas["cursor"]=="fleur" else None))
    canvas.bind("<B1-Motion>", lambda e: (canvas.scan_dragto(e.x, e.y, gain=1) if canvas["cursor"]=="fleur" else None))
    canvas.bind("<KeyRelease-space>", lambda e: canvas.configure(cursor=""))

    # Прокручивание колесиком мыши (только для этого canvas)
    def _on_mousewheel(e):
        # Проверяем, что событие от нашего canvas
        if e.widget != canvas:
            return
        delta = -1 if e.delta > 0 else 1
        canvas.yview_scroll(delta * (WHEEL_STEP // 30), "units")
        return "break"

    def _on_shift_mousewheel(e):
        # Проверяем, что событие от нашего canvas
        if e.widget != canvas:
            return
        delta = -1 if e.delta > 0 else 1
        canvas.xview_scroll(delta * (WHEEL_STEP // 30), "units")
        return "break"

    # Привязываем события только к canvas, а не ко всему приложению
    canvas.bind("<MouseWheel>", _on_mousewheel)
    canvas.bind("<Shift-MouseWheel>", _on_shift_mousewheel)
    
    # Linux-совместимость
    canvas.bind("<Button-4>", lambda e: (canvas.yview_scroll(-WHEEL_STEP // 30, "units"), "break"))
    canvas.bind("<Button-5>", lambda e: (canvas.yview_scroll(+WHEEL_STEP // 30, "units"), "break"))

    # ЗУМ: Ctrl + колёсико, Ctrl +/-, Ctrl + 0
    def _apply_zoom(new_scale: float):
        new_scale = max(SCALE_MIN, min(SCALE_MAX, new_scale))
        if abs(new_scale - SCALE["k"]) < 1e-3:
            return
        SCALE["k"] = new_scale
        build_all_matrices()
        # Обновляем метку масштаба
        try:
            scale_label.config(text=f"Масштаб: {int(SCALE['k'] * 100)}%")
        except:
            pass  # Если scale_label еще не создан

    def _on_ctrl_wheel(e):
        # Проверяем, что событие от нашего canvas
        if e.widget != canvas:
            return
        step = 1.1 if e.delta > 0 else 1/1.1
        _apply_zoom(SCALE["k"] * step)
        return "break"

    def _zoom_in(_=None):  
        _apply_zoom(SCALE["k"] * 1.1)
    def _zoom_out(_=None): 
        _apply_zoom(SCALE["k"] / 1.1)
    def _zoom_reset(_=None): 
        _apply_zoom(1.0)

    # Привязываем события только к canvas
    canvas.bind("<Control-MouseWheel>", _on_ctrl_wheel)
    canvas.bind("<Control-plus>",  _zoom_in)
    canvas.bind("<Control-equal>", _zoom_in)   # Ctrl + '='
    canvas.bind("<Control-minus>", _zoom_out)
    canvas.bind("<Control-0>",     _zoom_reset)
    
    # Добавляем клавиатурную навигацию
    def _on_key_press(event):
        if event.widget != canvas:
            return
        if event.keysym == 'Home':
            _nav_home()
        elif event.keysym == 'End':
            _nav_end()
        elif event.keysym == 'Left':
            _nav_left()
        elif event.keysym == 'Right':
            _nav_right()
        elif event.keysym == 'Up':
            _nav_up()
        elif event.keysym == 'Down':
            _nav_down()
        elif event.keysym == 'plus' or event.keysym == 'equal':
            _zoom_in()
        elif event.keysym == 'minus':
            _zoom_out()
        elif event.keysym == '0':
            _zoom_reset()
        elif event.keysym == 'c':
            canvas.xview_moveto(0.5)
            canvas.yview_moveto(0.5)
    
    canvas.bind('<KeyPress>', _on_key_press)
    canvas.focus_set()  # Делаем canvas фокусируемым

    canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    canvas.pack(side="left", fill="both", expand=True, padx=8, pady=(0,8))
    vsb.pack(side="right", fill="y"); hsb.pack(side="bottom", fill="x")
    
    # Добавляем обработку изменения размера canvas
    def _on_canvas_configure(event):
        # Обновляем scrollregion при изменении размера
        try:
            if hasattr(canvas, '_last_scrollregion'):
                canvas.configure(scrollregion=canvas._last_scrollregion)
        except:
            pass
    
    canvas.bind('<Configure>', _on_canvas_configure)
    
    # Сохраняем последний scrollregion для восстановления
    def _save_scrollregion():
        try:
            canvas._last_scrollregion = canvas.cget('scrollregion')
        except:
            pass
        # Метрики, зависящие от масштаба                      ### NEW
    def _metrics():
        k = SCALE["k"]
        CELL_W  = int(110 * k)
        CELL_H  = int(32  * k)
        PAD     = max(1, int(2 * k))
        HEADER_H= int(150 * k)
        V_GAP   = int(40  * k)
        LEFT_MIN= int(260 * k)
        LEFT_MAX= int(560 * k)
        return CELL_W, CELL_H, PAD, HEADER_H, V_GAP, LEFT_MIN, LEFT_MAX
    def build_one_matrix(line: str, y_offset: int) -> int:
        try:
            CELL_W, CELL_H, PAD, HEADER_H, V_GAP, LEFT_MIN, LEFT_MAX = _metrics()
            base_sz  = max(6, int(9  * SCALE["k"]))
            small_sz = max(6, int(8  * SCALE["k"]))
            big_sz   = max(7, int(10 * SCALE["k"]))
            fnt_base = tkFont.Font(family="Segoe UI", size=base_sz)
            fnt_bold = tkFont.Font(family="Segoe UI", size=base_sz, weight="bold")
            fnt_small_bold = tkFont.Font(family="Segoe UI", size=small_sz, weight="bold")
            fnt_big_bold   = tkFont.Font(family="Segoe UI", size=big_sz,   weight="bold")
        except Exception as e:
            print(f"Ошибка в метриках для линии {line}: {e}")
            return y_offset + 100

        try:
            sip_map = _build_sip_map_for_line(line)
            allow_pairs, deny_pairs = _build_evict_maps_for_line(line)
        except Exception as e:
            print(f"Ошибка при загрузке правил для линии {line}: {e}")
            canvas.create_text(
                20, y_offset + 20,
                text=f"Ошибка загрузки правил для линии «{line}»: {e}",
                anchor="nw", font=("Segoe UI", 10), fill="red"
            )
            return y_offset + 60
            
        if not sip_map:
            canvas.create_text(
                20, y_offset + 20,
                text=f"Для линии «{line}» нет назначенного набора или он пуст.",
                anchor="nw", font=("Segoe UI", 10)
            )
            return y_offset + 60
        norm_line = _low(line)
        rows = [r for r in rows_cache if _low(r.get("line", "")) == norm_line and r.get("name")]
        products_raw = sorted(set(((r.get("name", "") or ""), (r.get("container", "") or "")) for r in rows))
        products: list[tuple[str, str]] = []
        for nm, cont in products_raw:
            key = _prod_key_from_catalog(nm, cont)
            if key in sip_map:
                products.append((nm, cont))
        if not products:
            canvas.create_text(
                20, y_offset + 20,
                text=f"Линия «{line}»: ни один продукт каталога не попал в правила набора — матрица не построена.",
                anchor="nw", font=("Segoe UI", 10)
            )
            return y_offset + 60
        canvas.create_text(
            20, y_offset + 10,
            text=f"Линия: {line} • Правил: {len(sip_map)}",
            anchor="nw", font=("Segoe UI", 10, "italic"), fill="#666"
        )
        clear_product_parse_cache()
        metas = {p: parse_product_name(p[0], p[1]) for p in products}
        fnt = tkFont.Font(family="Segoe UI", size=9)
        max_len = max(fnt_base.measure(f"{p[0]} {p[1]}") for p in products)
        LEFT_W = max(LEFT_MIN, min(max_len + int(60 * SCALE["k"]), LEFT_MAX))

        for j, (p, v) in enumerate(products, 1):
            x = LEFT_W + (j - 0.5) * (CELL_W + PAD)
            text = f"{p}\n({v})" if v else p
            canvas.create_text(
            x, y_offset + HEADER_H - int(10 * SCALE["k"]),
            text=text, angle=60, anchor="sw", width=int(180 * SCALE["k"]),
            font=fnt_small_bold
        )

        y0 = y_offset + HEADER_H
        for i, (pf, vf) in enumerate(products, 1):
            y = (i + 0.5) * (CELL_H + PAD) + y0
            row_text = f"{pf}\n({vf})" if vf else pf
            key_from = _prod_key_from_catalog(pf, vf)
            base = sip_map.get(key_from, {}).get("base")
            base_tag = ""
            if base == "CIP1": base_tag = " [CIP1]"
            elif base == "CIP2": base_tag = " [CIP2]"
            elif base == "CIP3": base_tag = " [CIP3]"
            canvas.create_text(
                LEFT_W - int(8 * SCALE["k"]), y,
                text=row_text + (base_tag or ""), anchor="e",
                width=LEFT_W - int(30 * SCALE["k"]),
                font=fnt_bold
            )

            for j, (pt, vt) in enumerate(products, 1):
                x = LEFT_W + (j - 1) * (CELL_W + PAD)
                y1 = i * (CELL_H + PAD) + y0
                canvas.create_rectangle(x, y1, x + CELL_W, y1 + CELL_H, outline="#bbb")
                if i == j:
                    canvas.create_text(x + CELL_W/2, y1 + CELL_H/2, text="—", font=("Segoe UI", 10), fill="#777")
                    continue
                val_ps = _cell_mark(metas[(pf, vf)], metas[(pt, vt)])
                if val_ps == "П":
                    canvas.create_text(x + CELL_W/2, y1 + CELL_H/2, text="П", font=("Segoe UI", 10, "bold"), fill="#d97706")
                    continue
                elif val_ps == "С":
                    canvas.create_text(x + CELL_W/2, y1 + CELL_H/2, text="С", font=("Segoe UI", 10, "bold"), fill="#2563eb")
                    continue
                to_key = _prod_key_from_catalog(pt, vt)
                vol_from = _vol_to_ml(vf) or _vol_to_ml(metas[(pf,vf)].get("volume"))
                vol_to   = _vol_to_ml(vt) or _vol_to_ml(metas[(pt,vt)].get("volume"))
                same_volume = (vol_from is not None and vol_to is not None and vol_from == vol_to)
                if same_volume and ((key_from, to_key) in allow_pairs) and ((key_from, to_key) not in deny_pairs):
                    canvas.create_text(
                        x + CELL_W/2, y1 + CELL_H/2,
                        text="В", font=("Segoe UI", 10, "bold"), fill="#d97706"  # буква В = вытеснение
                    )
                    continue
                chosen = base or ""
                alts_map = (sip_map.get(key_from, {}) or {}).get("alts", {})
                for alt_cip in ("CIP1", "CIP2", "CIP3"):
                    if alt_cip == base:
                        continue
                    if to_key in (alts_map.get(alt_cip) or set()):
                        chosen = alt_cip
                        break
                cip_colors = {"CIP1": "#2563eb", "CIP2": "#1e9d52", "CIP3": "#7c3aed"}
                color = cip_colors.get(chosen, "#444")
                canvas.create_text(
                    x + CELL_W / 2, y1 + CELL_H / 2,
                    text=chosen, font=("Segoe UI", 10, "bold"), fill=color
                )
        try:
            height = (len(products) + 1) * (CELL_H + PAD) + HEADER_H
            return y_offset + height + V_GAP
        except Exception as e:
            print(f"Ошибка при завершении построения матрицы для линии {line}: {e}")
            return y_offset + 100
    def build_all_matrices():
        try:
            canvas.delete("all")
            y = 0
            chosen = sorted(SELECTED_LINES) if SELECTED_LINES else []
            if not chosen:
                canvas.create_text(20, 20, text="Линии не выбраны", anchor="nw", font=("Segoe UI",10))
                canvas.configure(scrollregion=(0, 0, 400, 100))
                return
            
            # Рассчитываем примерную ширину на основе количества продуктов
            max_products = 0
            for ln in chosen:
                norm_line = _low(ln)
                rows = [r for r in rows_cache if _low(r.get("line", "")) == norm_line and r.get("name")]
                products_raw = sorted(set(((r.get("name", "") or ""), (r.get("container", "") or "")) for r in rows))
                max_products = max(max_products, len(products_raw))
            
            # Рассчитываем ширину на основе метрик
            CELL_W, CELL_H, PAD, HEADER_H, V_GAP, LEFT_MIN, LEFT_MAX = _metrics()
            estimated_width = LEFT_MAX + max_products * (CELL_W + PAD) + 100
            
            for ln in chosen:
                y = build_one_matrix(ln, y)
            
            # Устанавливаем scrollregion с запасом
            final_width = max(1200, estimated_width, canvas.winfo_width() or 1200)
            final_height = max(600, y + 50)
            canvas.configure(scrollregion=(0, 0, final_width, final_height))
            
            # Сохраняем scrollregion для восстановления
            _save_scrollregion()
            
            # Принудительно обновляем canvas
            canvas.update_idletasks()
            
        except Exception as e:
            print(f"Ошибка при построении матриц: {e}")
            canvas.create_text(20, 20, text=f"Ошибка: {e}", anchor="nw", font=("Segoe UI",10), fill="red")
            canvas.configure(scrollregion=(0, 0, 400, 100))
    build_all_matrices()
    def _make_norms_tab(parent):
        tab = ttk.Frame(parent)
        bar = ttk.Frame(tab); bar.pack(fill="x", padx=8, pady=(8,4))
        btn_add  = ttk.Button(bar, text="Добавить"); btn_add.pack(side="left")
        btn_del  = ttk.Button(bar, text="Удалить");  btn_del.pack(side="left", padx=6)
        btn_save = ttk.Button(bar, text="Сохранить");btn_save.pack(side="left", padx=12)
        btn_load = ttk.Button(bar, text="Загрузить");btn_load.pack(side="left", padx=6)
        info = ttk.Label(bar, text="", foreground="#666"); info.pack(side="left", padx=12)
        line_cols = tuple(f"line{i}" for i in range(1, 11))
        columns   = ("category","event") + line_cols
        headers   = ("Категория","Event") + tuple(f"Линия {i}" for i in range(1, 11))
        wrap = ttk.Frame(tab); wrap.pack(fill="both", expand=True, padx=8, pady=(0,8))
        tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="extended")
        vs = ttk.Scrollbar(wrap, orient="vertical",   command=tree.yview)
        hs = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0,column=0,sticky="nsew"); vs.grid(row=0,column=1,sticky="ns"); hs.grid(row=1,column=0,sticky="ew")
        wrap.rowconfigure(0,weight=1); wrap.columnconfigure(0,weight=1)
        for cid,hdr in zip(columns,headers): tree.heading(cid, text=hdr)
        tree.column("category", width=240, anchor="w"); tree.column("event", width=160, anchor="w")
        for cid in line_cols: tree.column(cid, width=90, anchor="center")
        tree.tag_configure("multi", background="#fff59d")
        _edit_var = tk.StringVar()
        _edit_ent = None
        _edit_info = {"iid": None, "column": None}
        def _start_cell_edit(event):
            nonlocal _edit_ent
            if tree.identify("region", event.x, event.y) != "cell":
                return
            iid = tree.identify_row(event.y)
            colid = tree.identify_column(event.x)
            if not iid or not colid:
                return
            col_idx = int(colid[1:]) - 1
            colname = columns[col_idx]
            bbox = tree.bbox(iid, column=colname)
            if not bbox:
                return
            x, y, w, h = bbox
            cur = tree.set(iid, colname)
            _edit_var.set(cur)
            if _edit_ent is not None:
                _edit_ent.destroy()
            _edit_ent = ttk.Entry(tree, textvariable=_edit_var)
            _edit_ent.place(x=x + 1, y=y + 1, width=w - 2, height=h - 2)
            _edit_ent.focus_set()
            _edit_info["iid"] = iid
            _edit_info["column"] = colname
        def _commit_cell_edit(*_):
            nonlocal _edit_ent
            if _edit_ent is None: return
            iid=_edit_info["iid"]; colname=_edit_info["column"]; val=_edit_var.get().strip()
            tree.set(iid, colname, val); _apply_row_tags(iid)
            _edit_ent.destroy(); _edit_ent=None
        def _cancel_cell_edit(*_):
            nonlocal _edit_ent
            if _edit_ent is not None: _edit_ent.destroy(); _edit_ent=None
        tree.bind("<Double-1>", _start_cell_edit)
        tree.bind("<Return>",   lambda e: _commit_cell_edit())
        tree.bind("<Escape>",   lambda e: _cancel_cell_edit())
        tree.bind("<Button-1>", lambda e: (_commit_cell_edit(), None))
        def _apply_row_tags(iid: str):
            vals = tree.item(iid, "values"); line_cols_idx = range(2, 2+len(line_cols))
            any_multi = any((";" in (vals[idx] or "")) for idx in line_cols_idx)
            tree.item(iid, tags=("multi",) if any_multi else ())
        def _add_row():
            iid = tree.insert("", "end", values=("","","","","","","","","","","","")); _apply_row_tags(iid)
        def _del_rows():
            for iid in tree.selection(): tree.delete(iid)
        def _save():
            rows=[]
            for iid in tree.get_children(""):
                v=tree.item(iid,"values"); row={"category":v[0],"event":v[1]}
                for i,cid in enumerate(line_cols, start=2): row[cid]=v[i]
                rows.append(row)
            _save_json(_NORMS_JSON, rows); info.config(text=f"Сохранено: {os.path.basename(_NORMS_JSON)}")
        def _load():
            for iid in tree.get_children(""): tree.delete(iid)
            rows=_load_json(_NORMS_JSON, [])
            for r in rows:
                vals=[r.get("category",""), r.get("event","")] + [r.get(cid,"") for cid in line_cols]
                iid = tree.insert("", "end", values=tuple(vals)); _apply_row_tags(iid)
            info.config(text=f"Загружено: {len(rows)}")
        btn_add.configure(command=_add_row); btn_del.configure(command=_del_rows)
        btn_save.configure(command=_save); btn_load.configure(command=_load); _load()
        return tab
    nb.add(_make_norms_tab(nb), text="Нормативы")
    def _make_sets_tab(kind: str):
        is_rules = (kind == "rules")
        sets = _ensure_rules_sets() if is_rules else _ensure_evict_sets()
        tab = ttk.Frame(nb)
        top = ttk.Frame(tab); top.pack(fill="x", padx=8, pady=(8,4))
        ttk.Label(top, text="Набор:").pack(side="left")
        combo = ttk.Combobox(top, state="readonly", width=32)
        combo.pack(side="left", padx=(6,8))
        def _refresh_combo(select_id: str | None = None):
            nonlocal sets
            names = [f'{s.get("name","")}  ({", ".join(s.get("lines",[]) or []) or "без линий"})' for s in sets]
            combo["values"] = names
            if not sets:
                combo.set("")
                return
            idx = 0
            if select_id:
                for i,s in enumerate(sets):
                    if s["id"] == select_id:
                        idx = i; break
            combo.current(idx)
        def _current_set() -> dict | None:
            if not sets or combo.current() < 0: return None
            return sets[combo.current()]
        def _add_set():
            nonlocal sets
            s = {"id": uuid.uuid4().hex[:8], "name": "Новый набор", "lines": [], "rules": []}
            sets.append(s); _refresh_combo(s["id"]); _load_set_to_grid()
        def _rename_set():
            s = _current_set()
            if not s: return
            win = tk.Toplevel(tab); win.title("Переименовать набор")
            ttk.Label(win, text="Название:").pack(anchor="w", padx=8, pady=(8,4))
            e = ttk.Entry(win); e.insert(0, s.get("name","")); e.pack(fill="x", padx=8)
            def _ok():
                s["name"] = e.get().strip() or s["name"]
                _refresh_combo(s["id"]); win.destroy()
            ttk.Button(win, text="OK", command=_ok).pack(anchor="e", padx=8, pady=8)
        def _delete_set():
            nonlocal sets
            s = _current_set()
            if not s: return
            if not MB.askyesno("Удаление", f"Удалить набор «{s.get('name','')}»?"): return
            sid = s["id"]; sets = [x for x in sets if x["id"] != sid]
            if is_rules: _save_rules_sets(sets)
            else:        _save_evict_sets(sets)
            _refresh_combo(); _load_set_to_grid()
        ttk.Button(top, text="Новый", command=_add_set).pack(side="left")
        ttk.Button(top, text="Переименовать", command=_rename_set).pack(side="left", padx=4)
        ttk.Button(top, text="Удалить", command=_delete_set).pack(side="left")
        frmL = ttk.Frame(tab); frmL.pack(fill="x", padx=8, pady=(0,4))
        lbl_lines = ttk.Label(frmL, text="Линии: —"); lbl_lines.pack(side="left")
        def _assign_lines():
            nonlocal sets
            s = _current_set()
            if not s: return
            all_lines = sorted(set(lines))
            in_use = _lines_in_use(sets)
            win = tk.Toplevel(tab); win.title("Назначить линии набору")
            # Увеличим окно и разрешим изменение размеров
            try:
                px = tab.winfo_rootx() + 60; py = tab.winfo_rooty() + 60
                win.geometry(f"520x460+{px}+{py}")
            except Exception:
                win.geometry("520x460")
            win.minsize(420, 360); win.resizable(True, True)
            ttk.Label(win, text="Выберите линии (Ctrl/Shift):").pack(anchor="w", padx=8, pady=(8,4))
            lb = tk.Listbox(win, selectmode="extended")
            for ln in all_lines:
                mark = ""
                if ln in in_use and in_use[ln] != s["id"]:
                    mark = "  [занято]"
                lb.insert("end", ln + mark)
            # Добавим скроллбар для длинных списков
            wrap = ttk.Frame(win); wrap.pack(fill="both", expand=True, padx=8, pady=8)
            vsb = ttk.Scrollbar(wrap, orient="vertical", command=lb.yview)
            lb.configure(yscrollcommand=vsb.set)
            lb.pack(in_=wrap, side="left", fill="both", expand=True)
            vsb.pack(in_=wrap, side="right", fill="y")
            sel_idx = [i for i,ln in enumerate(all_lines) if ln in (s.get("lines") or [])]
            for i in sel_idx: lb.selection_set(i)
            def _ok():
                chosen = []
                for i, ln in enumerate(all_lines):
                    if ln in in_use and in_use[ln] != s["id"] and i in lb.curselection():
                        MB.showwarning("Конфликт", f"Линия «{ln}» уже привязана к другому набору.")
                        return
                    if i in lb.curselection():
                        chosen.append(ln)
                s["lines"] = chosen
                lbl_lines.config(text="Линии: " + (", ".join(chosen) if chosen else "—"))
                build_all_matrices()
                win.destroy()
            btns = ttk.Frame(win); btns.pack(fill="x", padx=8, pady=8)
            ttk.Button(btns, text="OK", command=_ok).pack(side="right")
        ttk.Button(frmL, text="Назначить линии", command=_assign_lines).pack(side="left", padx=8)
        if is_rules:
            cols=("product","CIP1","CIP2","CIP3")
            hdrs=("Продукт","CIP1","CIP2","CIP3")
            widths=(320,220,220,220)
        else:
            cols=("from","to","exceptions");  hdrs=("Из","В","Исключения"); widths=(320,320,900)
        grid, tree, btn_save, btn_load, info = _make_grid_tab(tab, cols, hdrs, widths)
        grid.pack(fill="both", expand=True)
        import tempfile
        def _load_set_to_grid():
            s = _current_set()
            lbl_lines.config(text="Линии: " + (", ".join(s.get("lines", [])) or "—") if s else "Линии: —")
            for iid in tree.get_children(""):
                tree.delete(iid)
            if not s:
                return
            rows = s.get("rules", [])
            try:
                with tempfile.NamedTemporaryFile("w+", suffix=".json", delete=False, encoding="utf-8") as tf:
                    json.dump(rows, tf, ensure_ascii=False, indent=2)
                    tf.flush()
                    grid._load_json(tf.name)
            except Exception as e:
                MB.showerror("Загрузка набора", f"Не удалось загрузить набор: {e}")
        def _save_sets():
            s = _current_set()
            if not s:
                return
            rows = []
            for iid in tree.get_children(""):
                raw = getattr(grid, "_raw_by_iid", None)
                if isinstance(raw, dict) and iid in raw:
                    row = {c: raw[iid].get(c, "") for c in cols}
                else:
                    row = {c: _unwrap(tree.set(iid, c)) for c in cols}
                rows.append(row)
            s["rules"] = rows
            if is_rules:
                cleaned, bad_rows = [], 0
                for r in (s.get("rules") or []):
                    parsed = _parse_sip_set_row(r)
                    if not parsed:
                        bad_rows += 1
                        continue
                    cleaned.append({
                        "product": r.get("product", ""),
                        "CIP1": r.get("CIP1", ""),
                        "CIP2": r.get("CIP2", ""),
                        "CIP3": r.get("CIP3", ""),
                    })
                s["rules"] = cleaned
                if bad_rows:
                    MB.showwarning("Правила", "Пропущено строк без корректной базы/формата: " f"{bad_rows}\n(Нужна ровно одна «База» в CIP1/2/3)")
            lines_map = {}
            for ss in sets:
                for ln in (ss.get("lines") or []):
                    if ln in lines_map and lines_map[ln] != ss["id"]:
                        MB.showwarning("Конфликт", f"Линия «{ln}» дублируется в нескольких наборах.")
                        return
                    lines_map[ln] = ss["id"]
            if is_rules:
                _save_rules_sets(sets)
            else:
                _save_evict_sets(sets)
            MB.showinfo("Сохранено", "Изменения сохранены.")
            build_all_matrices()
        def _validate_flavors():
            flv_keys = _collect_catalog_prodkeys(rows_cache)
            if is_rules:
                mism = []
                iids = list(tree.get_children(""))
                for rix, iid in enumerate(iids, start=1):
                    prod = (tree.set(iid, "product") or "").strip()
                    prod_key = _prod_key_from_rule(prod)
                    if prod and prod_key not in flv_keys:
                        mism.append({"row": rix, "column": "product", "value": prod, "issue": "нет в каталоге (тип+вкус)"})
                    for sip_col in ("CIP1","CIP2","CIP3"):
                        raw = (tree.set(iid, sip_col) or "")
                        if _low(raw) in ("база","base") or not raw.strip():
                            continue
                        for tok in [t.strip() for t in raw.split(";") if t.strip()]:
                            tok_key = _prod_key_from_rule(tok)
                            if not tok_key or tok_key not in flv_keys:
                                mism.append({"row": rix, "column": sip_col, "value": tok, "issue": "нет в каталоге (тип+вкус)"})
                _show_mismatch_report("Несоответствия: Правила (тип+вкус)", mism)
            else:
                mism = []
                iids = list(tree.get_children(""))

                for rix, iid in enumerate(iids, start=1):
                    f = (tree.set(iid, "from") or "").strip()
                    t_raw = (tree.set(iid, "to") or "").strip()
                    exc_raw = (tree.set(iid, "exceptions") or "")

                    f_key = _prod_key_from_rule(f)
                    if f and f_key not in flv_keys:
                        mism.append({"row": rix, "column": "from", "value": f, "issue": "нет в каталоге (тип+вкус)"})

                    # --- валидируем каждую цель «В» через ';' ---
                    if t_raw:
                        for tok in [tt.strip() for tt in t_raw.split(";") if tt.strip()]:
                            t_key = _prod_key_from_rule(tok)
                            if not t_key or t_key not in flv_keys:
                                mism.append({"row": rix, "column": "to", "value": tok, "issue": "нет в каталоге (тип+вкус)"})

                    if exc_raw.strip():
                        for tok in [tt.strip() for tt in exc_raw.split(";") if tt.strip()]:
                            k = _prod_key_from_rule(tok)
                            if not k or k not in flv_keys:
                                mism.append({"row": rix, "column": "exceptions", "value": tok, "issue": "нет в каталоге (тип+вкус)"})

                _show_mismatch_report("Несоответствия: Вытеснения (тип+вкус)", mism)

        def _reload_sets_file():
            nonlocal sets
            sets = _ensure_rules_sets() if is_rules else _ensure_evict_sets()
            _refresh_combo()
            _load_set_to_grid()
            build_all_matrices()
        btn_save.configure(command=_save_sets)
        btn_load.configure(command=_reload_sets_file)
        ttk.Button(top, text="Проверить вкусы (⛔)", command=_validate_flavors).pack(side="left", padx=8)
        combo.bind("<<ComboboxSelected>>", lambda _e: _load_set_to_grid())
        _refresh_combo()
        _load_set_to_grid()
        return tab
    nb.add(_make_sets_tab("rules"), text="Правила (наборы)")
    nb.add(_make_sets_tab("evict"), text="Вытеснения (наборы)")

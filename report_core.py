# -*- coding: utf-8 -*-
"""
report_core.py — расчётные функции без GUI:
- индексация простоев
- сводка (summary)
- компактный отчёт (план/факт/OEE/топ-3)
- OEE-матрица (для тепловой сетки)
Никаких tkinter и messagebox здесь нет.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional

import re
import math
import numbers

# ===== Индексы колонок блоков ===============================================
# Блок продуктов (products)
B1_COL_NAME = 0
B1_COL_BEG  = 1
B1_COL_END  = 2
B1_COL_DUR  = 3

# Блок простоев (downtimes) — ожидаем, что DESC «прицеплен» в конец
D2_COL_NAME   = 0
D2_COL_REASON = 1
D2_COL_KIND   = 2
D2_COL_BEG    = 3
D2_COL_END    = 4
D2_COL_MIN    = 5
D2_COL_DESC   = 6

# ===== Фильтры ===============================================================
@dataclass(frozen=True)
class FilterOpts:
    selected_lines: set[str]
    selected_days: set[str]
    current_line: str  # "Все" или конкретная "Линия NN" / "NN"

# ===== Утилиты ===============================================================
def _norm_name(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s.replace("\xa0", " "))
    s = s.replace("«", '"').replace("»", '"').replace("“", '"').replace("”", '"')
    return s

def _safe_minutes(x) -> int:
    try:
        if isinstance(x, str) and x.strip() == "":
            return 0
        return int(round(float(x)))
    except Exception:
        return 0

def _is_blank_time(x) -> bool:
    s = str(x).strip()
    return s in ("", "0", "00:00:00", "0:00:00")

def _to_float(x) -> Optional[float]:
    try:
        return float(str(x).replace(",", "."))
    except Exception:
        return None

def _row_speed_from_products(row: list) -> Optional[float]:
    # скорость ожидается в колонке E (index=4), если доступно
    if len(row) > 4:
        f = _to_float(row[4])
        if f is not None and f > 0:
            return f
    return None

def _extract_fact_qty(row: list) -> Optional[int]:
    try:
        for v in reversed(row):
            try:
                f = float(v)
                if math.isfinite(f):
                    return int(round(f))
            except Exception:
                continue
    except Exception:
        pass
    return None

def _is_planned(kind: str, reason: str = "") -> bool:
    def norm(s: str) -> str:
        return str(s or "").lower().replace("ё", "е")
    s = norm(kind) + " " + norm(reason)
    if "неплан" in s:
        return False
    return "план" in s

# ===== Индексация простоев ===================================================
def build_downtime_index(
    DATA: Dict[str, Dict],
    DOWNTIME_BLOCKS: List[str],
) -> Tuple[Dict[Tuple[str, str, str], List[List]], Dict[Tuple[str, str, str], Dict]]:
    """
    Возвращает:
      DOWNTIME_BY: (name_norm, day, shift) -> [events]
      AGG_BY:      (name_norm, day, shift) -> {"total_min":int, "events":int, "by_kind":dict, "top3":[(reason,mins),..]}
    """
    DOWNTIME_BY: Dict[Tuple[str, str, str], List[List]] = {}
    AGG_BY: Dict[Tuple[str, str, str], Dict] = {}

    def _has_meaning(r: List) -> bool:
        if not r:
            return False
        reason = str(r[D2_COL_REASON]).strip() if len(r) > D2_COL_REASON else ""
        kind   = str(r[D2_COL_KIND]).strip()   if len(r) > D2_COL_KIND   else ""
        beg    = str(r[D2_COL_BEG]).strip()    if len(r) > D2_COL_BEG    else ""
        end    = str(r[D2_COL_END]).strip()    if len(r) > D2_COL_END    else ""
        mins   = _safe_minutes(r[D2_COL_MIN] if len(r) > D2_COL_MIN else 0)
        return bool(reason or kind or beg or end or mins)

    for blk_name in DOWNTIME_BLOCKS:
        blk = DATA.get(blk_name)
        if not blk:
            continue
        rows = blk.get("array", [])
        meta = blk.get("meta", {})
        day   = str(meta.get("sheet", ""))
        shift = str(meta.get("shift", ""))

        if not rows:
            continue
        last_name = None
        for r in rows:
            if not r:
                continue
            cur = str(r[D2_COL_NAME]).strip() if len(r) > D2_COL_NAME else ""
            if cur not in ("", "0"):
                last_name = cur
            name = last_name
            if not name or not _has_meaning(r):
                continue
            key = (_norm_name(name), day, shift)
            DOWNTIME_BY.setdefault(key, []).append(r)

    for key, events in DOWNTIME_BY.items():
        total_min = 0
        kinds: Dict[str, int] = {}
        reasons: Dict[str, int] = {}
        for ev in events:
            m = _safe_minutes(ev[D2_COL_MIN] if len(ev) > D2_COL_MIN else 0)
            total_min += m
            kind = str(ev[D2_COL_KIND]).strip()   if len(ev) > D2_COL_KIND   else ""
            reason = str(ev[D2_COL_REASON]).strip() if len(ev) > D2_COL_REASON else ""
            if kind:
                kinds[kind] = kinds.get(kind, 0) + m
            if reason:
                reasons[reason] = reasons.get(reason, 0) + m
        top3 = sorted(reasons.items(), key=lambda kv: kv[1], reverse=True)[:3]
        AGG_BY[key] = {
            "total_min": int(total_min),
            "events": len(events),
            "by_kind": kinds,
            "top3": top3,
        }

    return DOWNTIME_BY, AGG_BY

# ===== Топ-3 для отчёта ======================================================
def top3_for(
    DOWNTIME_BY: Dict[Tuple[str, str, str], List[List]],
    name: str, day_label: str, shift_label: str
) -> List[dict]:
    key = (_norm_name(name), day_label, shift_label)
    events = DOWNTIME_BY.get(key, [])
    if not events:
        return []
    acc: Dict[str, dict] = {}
    for ev in events:
        mins   = _safe_minutes(ev[D2_COL_MIN] if len(ev) > D2_COL_MIN else 0)
        reason = str(ev[D2_COL_REASON]).strip() if len(ev) > D2_COL_REASON else ""
        kind   = str(ev[D2_COL_KIND]).strip()   if len(ev) > D2_COL_KIND   else ""
        desc   = str(ev[D2_COL_DESC]).strip()   if len(ev) > D2_COL_DESC   else ""
        if _is_planned(kind, reason):
            continue
        if not reason and mins <= 0 and not desc:
            continue
        if reason not in acc:
            acc[reason] = {"mins":0, "kind":"", "desc":""}
        acc[reason]["mins"] += mins
        if not acc[reason]["kind"] and kind:
            acc[reason]["kind"] = kind
        if not acc[reason]["desc"] and desc:
            acc[reason]["desc"] = desc
    items = [{"mins":v["mins"], "reason":r, "kind":v["kind"], "desc":v["desc"]} for r,v in acc.items()]
    items.sort(key=lambda x: x["mins"], reverse=True)
    return items[:3]

def fmt_top_item(item: dict) -> str:
    if not item:
        return ""
    mins = int(round(item.get("mins", 0)))
    reason = item.get("reason", "") or ""
    kind = item.get("kind", "") or ""
    desc = item.get("desc", "") or ""
    base = f"{mins} мин"
    if reason:
        base += f" • {reason}"
    if kind:
        base += f" [{kind}]"
    if desc:
        base += f" — {desc}"
    return base

# ===== Сводка по продуктам ===================================================
def build_summary_rows(
    DATA: Dict[str, Dict],
    PRODUCT_BLOCKS: List[str],
    DOWNTIME_BY: Dict[Tuple[str, str, str], List[List]],
    DEFAULT_SPEED_BY_LINE: Dict[str, float],
    flt: FilterOpts,
) -> Tuple[List[str], List[List]]:
    headers = [
        "Продукт","Линия","День","Смена",
        "Начало","Конец","Длит (мин)","Σ простой (мин)","% простоя",
        "Событий","План. простой (мин)","EffMin (мин)",
        "Ном. скорость (шт/ч)","Потолок (шт)","Факт (шт)","OEE, %",
    ]
    rows_out: List[List] = []

    for blk_name in PRODUCT_BLOCKS:
        blk = DATA.get(blk_name)
        if not blk:
            continue
        meta = blk.get("meta", {})
        day_label   = str(meta.get("sheet", ""))
        shift_label = str(meta.get("shift", ""))
        line_label  = str(meta.get("line", ""))

        # --- фильтрация по линиям и дням ---
        if flt.selected_days and day_label not in flt.selected_days:
           continue
    # если линия не указана в meta — пропускаем фильтр
        if line_label:
         if flt.selected_lines and line_label not in flt.selected_lines:
           continue
         if flt.current_line != "Все" and line_label != flt.current_line:
           continue


        for r in blk.get("array", []):
            if not r or len(r) <= B1_COL_NAME:
                continue
            name = str(r[B1_COL_NAME]).strip()
            if name in ("", "0"):
                continue
            beg = r[B1_COL_BEG] if len(r) > B1_COL_BEG else ""
            end = r[B1_COL_END] if len(r) > B1_COL_END else ""
            if _is_blank_time(beg) or _is_blank_time(end):
                continue

            run_min = _safe_minutes(r[B1_COL_DUR] if len(r) > B1_COL_DUR else 0)
            key = (_norm_name(name), day_label, shift_label)
            events = DOWNTIME_BY.get(key, [])

            total_dt = planned_dt = 0
            for ev in events:
                m = _safe_minutes(ev[D2_COL_MIN] if len(ev) > D2_COL_MIN else 0)
                total_dt += m
                reason = str(ev[D2_COL_REASON]).strip() if len(ev) > D2_COL_REASON else ""
                kind   = str(ev[D2_COL_KIND]).strip()   if len(ev) > D2_COL_KIND   else ""
                if _is_planned(kind, reason):
                    planned_dt += m

            pct_dt = (total_dt / run_min * 100.0) if run_min > 0 else 0.0
            eff_min = max(run_min - planned_dt, 0)

            speed = _row_speed_from_products(r)
            if speed is None:
                speed = DEFAULT_SPEED_BY_LINE.get((line_label or "").strip())
            cap_nom_qty = int(round(eff_min * speed / 60.0)) if (speed is not None) else None
            fact_qty = _extract_fact_qty(r)

            oee_str = ""
            if cap_nom_qty is not None and cap_nom_qty > 0 and fact_qty is not None:
                oee_pct = (fact_qty / float(cap_nom_qty)) * 100.0
                oee_str = f"{oee_pct:.1f}"

            rows_out.append([
                name,line_label,day_label,shift_label,beg,end,run_min,total_dt,
                round(pct_dt,1), int(len(events)), planned_dt, eff_min,
                (int(speed) if speed else ""), (cap_nom_qty if cap_nom_qty is not None else ""),
                (fact_qty if fact_qty is not None else ""), oee_str,
            ])
    return headers, rows_out

# ===== Компактный отчёт (для второй вкладки) =================================
def build_report_rows(
    DATA: Dict[str, Dict],
    PRODUCT_BLOCKS: List[str],
    DOWNTIME_BY: Dict[Tuple[str, str, str], List[List]],
    DEFAULT_SPEED_BY_LINE: Dict[str, float],
    flt: FilterOpts,
) -> Tuple[List[str], List[List]]:
    headers = ["Продукт","Линия","День","Смена","План, шт","Факт, шт","OEE, %","Топ-1","Топ-2","Топ-3"]
    rows: List[List] = []

    for blk_name in PRODUCT_BLOCKS:
        blk = DATA.get(blk_name)
        if not blk:
            continue
        meta = blk.get("meta", {})
        day_label   = str(meta.get("sheet", ""))
        shift_label = str(meta.get("shift", ""))
        line_label  = str(meta.get("line", ""))

        if flt.selected_lines and line_label not in flt.selected_lines:
            continue
        if flt.selected_days and day_label not in flt.selected_days:
            continue
        if flt.current_line != "Все" and line_label != flt.current_line:
            continue

        for r in blk.get("array", []):
            if not r or len(r) <= B1_COL_NAME:
                continue
            name = str(r[B1_COL_NAME]).strip()
            if not name or name == "0":
                continue
            beg = r[B1_COL_BEG] if len(r) > B1_COL_BEG else ""
            end = r[B1_COL_END] if len(r) > B1_COL_END else ""
            if _is_blank_time(beg) or _is_blank_time(end):
                continue

            run_min = _safe_minutes(r[B1_COL_DUR] if len(r) > B1_COL_DUR else 0)

            key_ev = (_norm_name(name), day_label, shift_label)
            events = DOWNTIME_BY.get(key_ev, [])
            planned_dt = 0
            for ev in events:
                m = _safe_minutes(ev[D2_COL_MIN] if len(ev) > D2_COL_MIN else 0)
                reason = str(ev[D2_COL_REASON]).strip() if len(ev) > D2_COL_REASON else ""
                kind   = str(ev[D2_COL_KIND]).strip()   if len(ev) > D2_COL_KIND   else ""
                if _is_planned(kind, reason):
                    planned_dt += m
            eff_min = max(run_min - planned_dt, 0)
            if eff_min <= 0:
                continue

            speed = _row_speed_from_products(r)
            if speed is None:
                speed = DEFAULT_SPEED_BY_LINE.get((line_label or "").strip())
            if speed is None or speed <= 0:
                continue

            plan_qty = int(round(eff_min * (speed / 60.0)))
            fact_qty = _extract_fact_qty(r)
            if fact_qty is None:
                continue

            oee_pct = (fact_qty / plan_qty * 100.0) if plan_qty > 0 else None

            top3 = top3_for(DOWNTIME_BY, name, day_label, shift_label)
            t1 = fmt_top_item(top3[0]) if len(top3) > 0 else ""
            t2 = fmt_top_item(top3[1]) if len(top3) > 1 else ""
            t3 = fmt_top_item(top3[2]) if len(top3) > 2 else ""

            rows.append([
                name, line_label, day_label, shift_label,
                plan_qty, int(fact_qty),
                (f"{oee_pct:.1f}" if (oee_pct is not None) else ""),
                t1, t2, t3
            ])
    return headers, rows

# ===== OEE-матрица ===========================================================
def compute_oee_matrix(
    DATA: Dict[str, Dict],
    PRODUCT_BLOCKS: List[str],
    DOWNTIME_BY: Dict[Tuple[str, str, str], List[List]],
    DEFAULT_SPEED_BY_LINE: Dict[str, float],
    flt: FilterOpts,
) -> Tuple[List[str], List[str], Dict[Tuple[str,str,str], Optional[float]], Dict[Tuple[str,str], Optional[float]], Dict[str, Optional[float]]]:
    fact_by: Dict[Tuple[str,str,str], float] = {}
    cap_by:  Dict[Tuple[str,str,str], float] = {}
    days_set, lines_set = set(), set()

    for blk_name in PRODUCT_BLOCKS:
        blk = DATA.get(blk_name)
        if not blk:
            continue
        meta = blk.get("meta", {})
        day_label   = str(meta.get("sheet", ""))
        shift_label = str(meta.get("shift", ""))
        line_label  = str(meta.get("line", ""))

        if flt.selected_lines and line_label not in flt.selected_lines:
            continue
        if flt.selected_days and day_label not in flt.selected_days:
            continue
        if flt.current_line != "Все" and line_label != flt.current_line:
            continue

        for r in blk.get("array", []):
            if not r or len(r) <= B1_COL_NAME:
                continue
            name = str(r[B1_COL_NAME]).strip()
            if not name or name == "0":
                continue
            beg = r[B1_COL_BEG] if len(r) > B1_COL_BEG else ""
            end = r[B1_COL_END] if len(r) > B1_COL_END else ""
            if _is_blank_time(beg) or _is_blank_time(end):
                continue

            run_min = _safe_minutes(r[B1_COL_DUR] if len(r) > B1_COL_DUR else 0)
            # плановые простои
            key_ev = (_norm_name(name), day_label, shift_label)
            events = DOWNTIME_BY.get(key_ev, [])
            planned_dt = 0
            for ev in events:
                m = _safe_minutes(ev[D2_COL_MIN] if len(ev) > D2_COL_MIN else 0)
                reason = str(ev[D2_COL_REASON]).strip() if len(ev) > D2_COL_REASON else ""
                kind   = str(ev[D2_COL_KIND]).strip()   if len(ev) > D2_COL_KIND   else ""
                if _is_planned(kind, reason):
                    planned_dt += m
            eff_min = max(run_min - planned_dt, 0)
            if eff_min <= 0:
                continue

            speed = _row_speed_from_products(r)
            if speed is None:
                speed = DEFAULT_SPEED_BY_LINE.get((line_label or "").strip())
            if speed is None or speed <= 0:
                continue

            cap  = eff_min * (speed / 60.0)
            fact = _extract_fact_qty(r)
            if fact is None:
                continue

            key = (day_label, line_label, shift_label)
            fact_by[key] = fact_by.get(key, 0.0) + float(fact)
            cap_by[key]  = cap_by.get(key, 0.0)  + float(cap)

            days_set.add(day_label)
            lines_set.add(line_label)

    # сортировки
    def _natural_key(s: str):
        parts = re.findall(r"\d+|\D+", str(s))
        out=[]
        for p in parts:
            if p.isdigit():
                out.append((0,int(p)))
            else:
                out.append((1,p.lower()))
        return tuple(out)

    days_sorted  = sorted(days_set, key=lambda s: int(s)) if days_set else []
    lines_sorted = sorted(lines_set, key=_natural_key)

    # клетки
    cell: Dict[Tuple[str,str,str], Optional[float]] = {}
    for day in days_sorted:
        for line in lines_sorted:
            for shift in ("День","Ночь"):
                k=(day,line,shift)
                f=fact_by.get(k,0.0)
                c=cap_by.get(k,0.0)
                cell[k]=(f/c*100.0) if c>0 else None

    # итоги
    totals_shift: Dict[Tuple[str,str], Optional[float]] = {}
    totals_line:  Dict[str, Optional[float]] = {}

    for line in lines_sorted:
        # по сменам
        for shift in ("День","Ночь"):
            f_sum=c_sum=0.0
            for day in days_sorted:
                k=(day,line,shift)
                f_sum+=fact_by.get(k,0.0)
                c_sum+=cap_by.get(k,0.0)
            totals_shift[(line,shift)] = (f_sum/c_sum*100.0) if c_sum>0 else None
        # общий по линии
        f_sum=c_sum=0.0
        for shift in ("День","Ночь"):
            for day in days_sorted:
                k=(day,line,shift)
                f_sum+=fact_by.get(k,0.0)
                c_sum+=cap_by.get(k,0.0)
        totals_line[line]=(f_sum/c_sum*100.0) if c_sum>0 else None

    return days_sorted, lines_sorted, cell, totals_shift, totals_line

# -*- coding: utf-8 -*-
# product_parse.py — единый парсер наименований для всего проекта
from __future__ import annotations
import re
from functools import lru_cache
from typing import Dict, Iterable, Optional, Tuple, List

# ========================= ПАТТЕРНЫ И СЛОВАРИ =================================

# Тип продукта
_TYPE_RX = re.compile(
    r"\b(концентрат|сироп|основа|нектар|сок|напиток|морс|топпинг|соус)\b",
    re.I,
)

# Конструкция «ТМ … / TM …» с любыми кавычками/скобками
_TM_TAIL_RX = re.compile(
    r"(?:\bТМ\b|\bTM\b|\bтм\b|\btm\b)\s*[«\"'(].*?[»\")]?",
    re.I,
)

# Кавычки/ёлочки для зачистки
_QUOTES_RX = re.compile(r"[«»\"“”„]+")

# Неразрывные пробелы → обычные
_NBSP_RX = re.compile(r"\u00A0")

# Хвосты "… шт" и токен объёма "… л" — чтобы чистить имя перед извлечением вкуса
_VOL_TOKEN_RX = re.compile(r"\b\d+(?:[.,]\d+)?\s*л\b", re.I)
_QTY_TAIL_RX  = re.compile(r"[-–—]?\s*\d[\d\s]*\s*шт\.?\s*$", re.I)
# ========================= ИСКЛЮЧЕНИЯ =========================================

# Брендовые ключевые слова (линейки, детские, полезные и т.д.)
_BRAND_KEYWORDS = [
    "дети", "детский", "полезный", "100% джус", "джус", "juicer"
]

# Слова упаковки (не должны попадать во вкус)
_CONTAINER_WORDS = ["пэт", "pet", "пэт-бутылка", "бутылка", "тара", "пластик"]


# ========================= БРЕНДЫ ============================================

_BRAND_CANON: Tuple[Tuple[re.Pattern, str], ...] = (
    # Основные бренды
    (re.compile(r"(?:barinoff|бариноф{1,2}|бариновф)", re.I), "Баринофф"),
    (re.compile(r"rioba", re.I), "RIOBA"),
    (re.compile(r"gold\s*label", re.I), "Gold Label"),
    (re.compile(r"(?:auchan|ашан)", re.I), "Ашан"),
    (re.compile(r"(?:o'?key|окей)", re.I), "Окей"),
    (re.compile(r"100%\s*джус", re.I), "100% Джус"),
    (re.compile(r"spar\s*kids?", re.I), "Spar Kids"),
    (re.compile(r"\bspar\b", re.I), "Spar"),
    (re.compile(r"olife|олайф", re.I), "OLIFE"),
    (re.compile(r"baby\s*island", re.I), "Baby Island"),
    (re.compile(r"(?:вяу|вау\s*мяу|вяу\s*мяу|wow\s*meow)", re.I), "Вау Мяу"),
    (re.compile(r"тигруля", re.I), "Тигруля"),
    (re.compile(r"мамина\s*дача", re.I), "Мамина Дача"),
    (re.compile(r"для\s*всей\s*семьи", re.I), "Для Всей Семьи"),
    (re.compile(r"магнит", re.I), "Магнит"),
    (re.compile(r"миш[аы]", re.I), "Миша"),
    (re.compile(r"динозаврик\s*ди", re.I), "Динозаврик ДИ"),
    (re.compile(r"русский\s*морс", re.I), "Русский Морс"),
    (re.compile(r"granatel|гранател", re.I), "Granatel"),
    (re.compile(r"гранатовый\s*рай", re.I), "Гранатовый Рай"),
    (re.compile(r"soko\s*grande", re.I), "Soko Grande"),
    (re.compile(r"каждый\s*день", re.I), "Каждый День"),
    (re.compile(r"365\s*дней", re.I), "365 Дней"),
    (re.compile(r"полезный\s*сок", re.I), "Полезный Сок"),
    (re.compile(r"додо", re.I), "Додо"),
)

# Приоритет, если встречается несколько брендов
_BRAND_PRIORITY: Tuple[str, ...] = (
    "Gold Label",
    "RIOBA",
    "Ашан",
    "Окей",
    "Spar Kids",
    "Spar",
    "100% Джус",
    "Baby Island",
    "OLIFE",
    "Granatel",
    "Гранатовый Рай",
    "Русский Морс",
    "Вау Мяу",
    "Миша",
    "Тигруля",
    "Мамина Дача",
    "Для Всей Семьи",
    "Магнит",
    "Динозаврик ДИ",
    "Полезный Сок",
    "Баринофф",
    "Додо",
)

# Регекс для вырезания брендов из названия вкуса
_BRAND_TOKEN_RX = re.compile(
    r'(?:^|[\s,;:-])'
    r'[«"""\u201e\']?'
    r'(?:barinoff|бариноф{1,2}|бариновф|'
    r'rioba|gold\s*label|'
    r'100%\s*джус|'
    r'o\'?key|окей|'
    r'spar\s*kids?|\bspar\b|'
    r'olife|олайф|'
    r'baby\s*island|'
    r'vau\s*meow|вяу|вау\s*мяу|вяу\s*мяу|'
    r'тигруля|'
    r'мамина\s*дача|'
    r'для\s*всей\s*семьи|'
    r'магнит|'
    r'миш[аы]|'
    r'динозаврик\s*ди|'
    r'русский\s*морс|'
    r'granatel|гранатовый\s*рай|'
    r'soko\s*grande|'
    r'каждый\s*день|365\s*дней|'
    r'полезный\s*сок|'
    r'ашан|auchan|'
    r'додо)'
    r'[»"""\u201e\']?'
    r'(?=$|[\s,;:\u2013\u2014-])',
    re.I
)

# ========================= УТИЛИТЫ НОРМАЛИЗАЦИИ ===============================

def _norm(s: object) -> str:
    return _NBSP_RX.sub(" ", str(s or "")).strip()

def _low(s: str) -> str:
    return _norm(s).lower().replace("ё", "е")

def _strip_tm(src: str) -> str:
    """Удаляет хвост 'ТМ …' / 'TM …' вместе с кавычками/скобками и чистит края."""
    s = _TM_TAIL_RX.sub("", src)
    s = _QUOTES_RX.sub("", s)
    s = s.strip(" ,;:-—·•'")
    s = re.sub(r"\s{2,}", " ", s)
    return s

def _canon_brand(low_str: str) -> str:
    """Определяет бренд по паттернам и приоритетам."""
    found: List[str] = []
    for rx, canon in _BRAND_CANON:
        if rx.search(low_str):
            found.append(canon)
    if not found:
        return ""
    for b in _BRAND_PRIORITY:
        if b in found:
            return b
    return found[0]

def _strip_brand_tokens(s: str) -> str:
    """Удаляет любые упоминания брендов из середины/конца строки вкуса."""
    s = _BRAND_TOKEN_RX.sub(" ", _low(s))
    s = re.sub(r"\s{2,}", " ", s).strip(" ,;:-—·•'")
    # вернём капс для кириллицы красиво
    s = " ".join(w.capitalize() if w.isalpha() else w for w in s.split())
    return s

# =============================== ПАРСЕР =======================================

@lru_cache(maxsize=8192)
def parse_product_name(name: str, volume: str = "") -> Dict[str, str]:
    """
    Возвращает словарь:
      {'type': str, 'flavor': str, 'brand': str, 'volume': str}

    Логика:
    - аккуратно чистим из name объём (например, '1,0 л') и хвост '... шт' — чтобы не попадали во flavor;
    - вытаскиваем type (категорию) и brand;
    - flavor = name без типа/ТМ/брендов/служебных хвостов.
    """
    # исходник + базовая нормализация пробелов
    src = _norm(name)
    low = _low(src)

    # --- предочистка имени от объёма и количества (защита flavor)
    # только шаблоны 'число л' и '... шт' в КОНЦЕ строки
    src = _QTY_TAIL_RX.sub("", src)   # убираем " - 18 000 шт"
    src = _VOL_TOKEN_RX.sub("", src)  # убираем "1,0 л", "0.33 л"
    src = re.sub(r"\s{2,}", " ", src).strip(" ,;:-—·•'")
    low = _low(src)  # обновлённая нижняя строка

    # тип
    m_type = _TYPE_RX.search(low)
    ptype = (m_type.group(1) if m_type else "").lower()

    # бренд
    brand = _canon_brand(low)
        # Если бренд не найден через паттерны, ищем по ключевым словам
    if not brand:
        for b in _BRAND_KEYWORDS:
            if re.search(rf"\b{b}\b", low, re.I):
                brand = b.capitalize()
                src = re.sub(rf"\b{b}\b", "", src, flags=re.I)
                break


    # вкус: удаляем тип, ТМ и брендовые токены
        # --- ВКУС ---
    flavor = src
    if ptype:
        flavor = re.sub(rf"\b{re.escape(ptype)}\b", "", flavor, flags=re.I)
    flavor = _strip_tm(flavor)
    flavor = _strip_brand_tokens(flavor)
    flavor = _QUOTES_RX.sub("", flavor).strip(" ,;:-—·•'")
    flavor = re.sub(r"\s{2,}", " ", flavor)

    # Убираем упаковку и бренды из вкуса
    for bad in _BRAND_KEYWORDS + _CONTAINER_WORDS:
        flavor = re.sub(rf"\b{bad}\b", "", flavor, flags=re.I)
    flavor = re.sub(r"\s{2,}", " ", flavor).strip(" ,;:-—·•'")


    return {
        "type": ptype,
        "flavor": flavor,
        "brand": brand,
        "volume": _norm(volume),
    }

# =============================== УТИЛИТЫ ======================================

def parse_pairs(pairs: Iterable[Tuple[str, str]]) -> List[Dict[str, str]]:
    """(name, volume) → список словарей парсера."""
    return [parse_product_name(name, vol) for name, vol in pairs]

def parse_catalog(
    catalog: object,
    line: Optional[str] = None,
    name_field: str = "name",
    volume_field: str = "container",
) -> List[Dict[str, str]]:
    """
    Из объекта с методом rows() берём пары (name, volume) и парсим.
    line=None — без фильтра по линии.
    """
    rows = catalog.rows()
    if line is not None:
        rows = [r for r in rows if r.get("line") == line]
    pairs = [(r.get(name_field, ""), r.get(volume_field, "")) for r in rows if r.get(name_field)]
    return parse_pairs(pairs)

def clear_product_parse_cache() -> None:
    """Сбрасывает LRU-кэш парсера."""
    try:
        parse_product_name.cache_clear()
    except AttributeError:
        pass

__all__ = ["parse_product_name", "parse_pairs", "parse_catalog", "clear_product_parse_cache"]

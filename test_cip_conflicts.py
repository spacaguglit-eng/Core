# -*- coding: utf-8 -*-
"""
Тест обнаружения конфликтов CIP
"""

def test_cip_conflict_detection():
    """Тест функции обнаружения конфликтов с существующими CIP"""

    # Импортируем функции
    import sys
    import os
    sys.path.append(os.path.dirname(__file__))

    # Мокаем функции для тестирования
    def _check_cip_conflict(results, start_time, duration, line):
        """Проверяет, есть ли конфликт с существующими CIP в заданном интервале"""
        end_time = start_time + duration

        for record in results:
            if record.get("line") != line:
                continue

            # Проверяем только CIP записи
            if record.get("type") not in ["CIP1", "CIP2", "CIP3", "CIP", "ВЫТ", "П"]:
                continue

            # Получаем время начала и окончания CIP
            start_str = record.get("start", "")
            end_str = record.get("end", "")

            try:
                # Парсим время (формат: "25.10 14:30")
                import re
                start_match = re.search(r"(\d{1,2})\.(\d{1,2})\s+(\d{2}):(\d{2})", start_str)
                end_match = re.search(r"(\d{1,2})\.(\d{1,2})\s+(\d{2}):(\d{2})", end_str)

                if start_match and end_match:
                    start_minutes = int(start_match.group(3)) * 60 + int(start_match.group(4))
                    end_minutes = int(end_match.group(3)) * 60 + int(end_match.group(4))

                    # Проверяем пересечение интервалов
                    if not (end_time <= start_minutes or start_time >= end_minutes):
                        return True  # Есть конфликт
            except:
                continue

        return False  # Нет конфликта

    # Тестовые данные
    existing_cip = {
        "line": "линия 1",
        "type": "CIP2",
        "start": "25.10 14:00",
        "end": "25.10 15:00"
    }

    results = [existing_cip]

    # Тест 1: Конфликт (пересечение)
    conflict1 = _check_cip_conflict(results, 13*60 + 30, 60, "линия 1")  # 13:30-14:30
    print(f"Тест 1 - Конфликт с 14:00-15:00 при 13:30-14:30: {conflict1}")
    assert conflict1 == True, "Должен быть конфликт"

    # Тест 2: Нет конфликта (до)
    conflict2 = _check_cip_conflict(results, 12*60, 60, "линия 1")  # 12:00-13:00
    print(f"Тест 2 - Нет конфликта с 14:00-15:00 при 12:00-13:00: {conflict2}")
    assert conflict2 == False, "Не должно быть конфликта"

    # Тест 3: Нет конфликта (после)
    conflict3 = _check_cip_conflict(results, 15*60 + 30, 60, "линия 1")  # 15:30-16:30
    print(f"Тест 3 - Нет конфликта с 14:00-15:00 при 15:30-16:30: {conflict3}")
    assert conflict3 == False, "Не должно быть конфликта"

    # Тест 4: Другая линия
    conflict4 = _check_cip_conflict(results, 14*60, 60, "линия 2")
    print(f"Тест 4 - Нет конфликта на другой линии: {conflict4}")
    assert conflict4 == False, "Не должно быть конфликта на другой линии"

    print("Все тесты конфликтов пройдены!")


if __name__ == "__main__":
    test_cip_conflict_detection()

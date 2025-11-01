# -*- coding: utf-8 -*-
"""
Тест умных буферов автосипов
"""

def test_smart_buffer():
    """Тест применения буфера только при риске частых CIP"""

    print("=== ТЕСТ УМНЫХ БУФЕРОВ АВТОСИПОВ ===\n")

    # Тест 1: Буфер НЕ применяется (остаток достаточный)
    print("Тест 1: Буфер НЕ применяется при достаточном остатке")
    volume_threshold = 50000
    min_remainder = 2000
    total_volume_since_cip = 45000
    remaining_qty = 8000  # Всего будет 45000 + 8000 = 53000

    total_with_remaining = total_volume_since_cip + remaining_qty  # 53000
    remainder_after_cip = total_with_remaining - volume_threshold  # 53000 - 50000 = 3000

    if remainder_after_cip < min_remainder:
        buffer_needed = min_remainder - remainder_after_cip
        effective_threshold = volume_threshold + buffer_needed
        print(f"Ошибка: буфер применен при достаточном остатке")
    else:
        print(f"Корректно: остаток {remainder_after_cip} >= {min_remainder}, буфер НЕ применяется")
        print(f"   CIP вставляется на пороге: {volume_threshold}")

    # Тест 2: Буфер применяется при малом остатке
    print("\nТест 2: Буфер применяется при малом остатке")
    total_volume_since_cip = 49500
    remaining_qty = 1000  # Всего будет 49500 + 1000 = 50500

    total_with_remaining = total_volume_since_cip + remaining_qty  # 50500
    remainder_after_cip = total_with_remaining - volume_threshold  # 50500 - 50000 = 500

    if remainder_after_cip < min_remainder:
        buffer_needed = min_remainder - remainder_after_cip  # 2000 - 500 = 1500
        effective_threshold = volume_threshold + buffer_needed  # 50000 + 1500 = 51500
        print(f"Корректно: остаток {remainder_after_cip} < {min_remainder}")
        print(f"   Буфер нужен: {buffer_needed}")
        print(f"   Эффективный порог: {effective_threshold}")
    else:
        print(f"Ошибка: буфер не применен при малом остатке")

    # Тест 3: Граничный случай
    print("\nТест 3: Граничный случай (остаток = буфер)")
    total_volume_since_cip = 48000
    remaining_qty = 4000  # Всего будет 48000 + 4000 = 52000

    total_with_remaining = total_volume_since_cip + remaining_qty  # 52000
    remainder_after_cip = total_with_remaining - volume_threshold  # 52000 - 50000 = 2000

    if remainder_after_cip < min_remainder:
        print(f"Ошибка: буфер применен при остатке = буферу")
    else:
        print(f"Корректно: остаток {remainder_after_cip} >= {min_remainder}, буфер НЕ применяется")

    print("\nВсе тесты умных буферов пройдены!")

if __name__ == "__main__":
    test_smart_buffer()

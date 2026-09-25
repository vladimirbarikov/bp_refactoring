#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
# pylint: disable=import-outside-toplevel
"""
BP Refactoring Tool - Пошаговая обработка технических изменений Breakpoint (BP) автокомпонентов.

Модуль является чистой библиотекой пайплайна обработки одного BP-файла.
Точка входа приложения (оркестрация сессии) находится в корневом main.py,
который импортирует из этого модуля функцию process_bp_file().

Основная функция модуля:
    process_bp_file(bp_filename) -> Optional[Dict[str, Any]]

Она выполняет 25 последовательных шагов обработки одного BP<номер>.xlsx:
    ШАГ 1: ЗАГРУЗКА, АУДИТ И НОРМАЛИЗАЦИЯ ДАННЫХ (BP<номер>.xlsx)
    ШАГ 2: РАЗДЕЛЕНИЕ ДЕТАЛЕЙ НА СТАРЫЕ И НОВЫЕ (BEFORE / AFTER)
    ШАГ 3: ОПРЕДЕЛЕНИЕ СТАТУСА (Status)
    ШАГ 4: ОПРЕДЕЛЕНИЕ НАЗВАНИЙ МОДЕЛЕЙ ПО КОДАМ (BOM Product Code -> BOM Product Name)
    ШАГ 5: ВВОД ФАКТИЧЕСКОЙ ПАРТИИ (Batch Fact)
    ШАГ 6: ВВОД ФАКТИЧЕСКОЙ ДАТЫ (Change Date)
    ШАГ 7: ВВОД КОЛИЧЕСТВА BEFORE-ДЕТАЛЕЙ В SAFETY STOCK (Quantity Parts in SS)
    ШАГ 8: ПЕРЕВОД НАЗВАНИЙ ДЕТАЛЕЙ (Part Name Before / After)
    ШАГ 9: ПЕРЕВОД ОФИЦИАЛЬНЫХ НАЗВАНИЙ ПОСТАВЩИКОВ (Supplier Name Before / After)
    ШАГ 10: ВВОД СТАТУСА ЛОКАЛИЗАЦИИ ПОСТАВЩИКОВ (Localization Before / After)
    ШАГ 11: ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ BEFORE-ПОСТАВЩИКОВ
    ШАГ 12: ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ AFTER-ПОСТАВЩИКОВ
    ШАГ 13: ПЕРЕВОД ОПИСАНИЯ К ИЗМЕНЕНИЮ (Change Description)
    ШАГ 14: ПЕРЕВОД РЕШЕНИЯ К ИЗМЕНЕНИЮ (Change Solution)
    ШАГ 15: ПЕРЕВОД ЦВЕТА (Color Code и Color Name)
    ШАГ 16: ОБРАБОТКА ЦЕХОВ И РАБОЧИХ СТАНЦИЙ (Workshop/Workcenter Code/Name Before/After)
    ШАГ 17: ПЕРЕВОД ТРЕБОВАНИЙ ПО УТИЛИЗАЦИИ СТАРЫХ ДЕТАЛЕЙ (Production Part Disposal)
    ШАГ 18: ПЕРЕВОД ТРЕБОВАНИЙ ПО ВЗАИМОЗАМЕНЯЕМОСТИ (Interchangeable)
    ШАГ 19: ЗАГРУЗКА КОНФИГУРАЦИОННОГО ФАЙЛА (YYYY-mm-dd_configuration.xlsx)
    ШАГ 20: РАСЧЁТ ПАРТИЙ В SAFETY STOCK (Quantity Batches in SS)
    ШАГ 21: ПОИСК КОНФИГУРАЦИИ ДЛЯ УТИЛИЗАЦИИ BEFORE-ДЕТАЛЕЙ
    ШАГ 22: ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ BEFORE-ДЕТАЛЕЙ
    ШАГ 23: ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ AFTER-ДЕТАЛЕЙ
    ШАГ 24: УПОРЯДОЧИВАНИЕ КОЛОНОК
    ШАГ 25: СОХРАНЕНИЕ РЕЗУЛЬТАТА

Каждый шаг включает:
    - Интерактивное взаимодействие с пользователем для ввода данных
    - Возможность отменить изменения и повторить шаг (режим retry)
    - Сохранение состояния перед каждым шагом для возможности восстановления

Экспортируемые функции:
    - process_bp_file()                  — полный 25-шаговый пайплайн одного BP
    - load_configuration_file()          — ленивая загрузка configuration.xlsx (кэш)
    - get_bp_status()                    — интерактивный ввод статуса
    - get_batch_fact()                   — интерактивный ввод Batch Fact
    - get_change_date()                  — интерактивный ввод Change Date
    - get_quantity_in_ss()               — интерактивный ввод остатков в Safety Stock
    - get_translation()                  — интерактивный ввод переводов
    - get_supplier_localization_status() — интерактивный ввод локализации
    - get_supplier_location_data()       — интерактивный ввод местоположения
    - detect_workshop_by_code()          — определение цеха по коду воркцентра
    - load_packing_list_file()           — загрузка упаковочного листа
    - parse_dimensions()                 — парсинг строки размера
    - collect_unique_parts()             — сбор уникальных деталей из матрицы
    - calculate_packaging_metrics()      — расчёт производных упаковочных метрик
    - get_packaging_user_inputs()        — интерактивный ввод типа упаковки
    - apply_packaging_to_rows()          — заполнение 18 упаковочных колонок
    - copy_packaging_before_to_after()   — автокопирование упаковки Before → After

Использование:
    from py_lib.pipeline.processing import process_bp_file
    result = process_bp_file('input_files/input_breakpoint_files/BP26000001.xlsx')

Версия: 2.0
Совместимость: Python 3.14.4+, Pandas 3.0.3+, OpenPyXL 3.1.5+
Поддержка: PLD Engineering Center
Дата создания: 2026-05-14
Дата изменения: 2026-09-24
Лицензия: MIT
Статус: Production
"""

import os
import re
import time
import warnings
from datetime import timedelta
from functools import lru_cache
from typing import Any, Dict, Hashable, List, Optional, Tuple

import numpy as np
import pandas as pd

from py_lib.config.core import (
    # Пути к директориям
    # --- Входные данные ---
    INPUT_CONFIGURATION_DIR,
    INPUT_PACKING_LIST_DIR,

    # Префиксы файлов входных/выходных данных
    CONFIGURATION_PREFIX,

    # Колонки с полными данными
    BP_GMOM_TO_KEEP_COLS,
    CONF_TO_KEEP_COLS,

    # Колонки по типам данных
    BP_GBOM_INT_COLS,
    BP_GBOM_DATETIME_COLS,

    # Колонки по типу Before/After данных упаковки
    BEFORE_PAC_COLS, AFTER_PAC_COLS,

    # Матрица определения цехов
    WORKSHOPS, WORKSHOP_PREFIX_TO_CODE,

    # Матрица определения моделей
    BOM_PRODUCT_MAP,

    # Массив колонок предпоказа результирующих данных
    STEP_FOCUS_COLUMNS_MAP,

    # Список порядка колонок в итоговом датафрейме
    BP_DATA_COLUMNS_ORDER,

    # Типы упаковки
    PACKAGING_TYPE
)

from py_lib.engine.etl import (
    find_latest_excel_file,
    read_excel_file,
    data_absence_marker,
    normalize_data,
    get_unique_non_empty_values,
    keep_chinese_lines,
    extract_parentheses_content,
    extract_packaging_data,
)

from py_lib.pipeline.find_pairs import find_pairs

from py_lib.ui.interaction import (
    ask_int_or_nan,
    ask_user_input,
    wait_for_user,
    print_step_header,
    show_dataframe_preview,
    save_state,
    confirm_step,
)

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def get_bp_status(
    bp_number: str
) -> str:
    """
    Запрашивает у пользователя статус для BP файла.

    Доступные варианты:
    1. Согласован
    2. Опубликован
    3. Закрыт
    4. Другое (ввод кастомного статуса вручную)

    Аргументы:
        bp_number (str): Номер технического изменения для вывода в сообщении.

    Возвращается:
        str: Выбранный статус на русском языке.
    """
    print(f"\n  Для BP файла {bp_number} укажите статус обработки:")
    print("  Доступные варианты:")
    print("    1 - Согласован")
    print("    2 - Опубликован")
    print("    3 - Закрыт")
    print("    4 - Другое (ввести вручную)")

    while True:
        choice = ask_user_input("  Ваш выбор (1-4, или Enter для 'Согласован'): ").strip()

        if choice == '' or choice == '1':
            return "Согласован"
        elif choice == '2':
            return "Опубликован"
        elif choice == '3':
            return "Закрыт"
        elif choice == '4':
            custom_status = ask_user_input("  Введите статус на русском языке: ").strip()

            if custom_status:
                return custom_status

            return "Согласован"
        else:
            print("  [Ошибка] Неверный выбор. Пожалуйста, введите число от 1 до 4.")


def get_batch_fact(
    bp_number: str
) -> str:
    """
    Запрашивает у пользователя номер фактической партии (Batch fact) 
    для текущего технического изменения (BP).

    Аргументы:
        bp_number (str): Номер технического изменения (например, 'BP25010410').

    Возвращается:
        str: Введенный номер фактической партии или пустая строка.
    """
    print(f"\n  [Ввод данных] Настройка партии для {bp_number}")

    # Исполняем чистый защищенный ввод без KeyboardInterrupt
    batch_fact = ask_user_input("  Введите номер фактической партии (Batch fact) [Enter — оставить пустым]: ")

    if batch_fact:
        print(f"    ✓ Зафиксировано Batch fact: '{batch_fact}'")
    else:
        print("    • Поле Batch fact оставлено пустым (можно заполнить позже в Excel)")

    return batch_fact


def get_change_date(
    bp_number: str
) -> Optional[str]:
    """
    Запрашивает у пользователя фактическую дату внедрения (Change Date) 
    с обязательной валидацией инженерного формата ГГГГ-ММ-ДД.

    Аргументы:
        bp_number (str): Номер технического изменения (например, 'BP25010410').

    Возвращается:
        Optional[str]: Валидная строка даты в формате YYYY-MM-DD или None.
    """
    print(f"\n  [Ввод данных] Настройка даты для {bp_number}")

    while True:
        change_date = ask_user_input("  Введите дату изменения (Change Date) в формате ГГГГ-ММ-ДД [Enter — пропустить]: ")

        if not change_date:
            print("    • Поле Change Date оставлено пустым (можно заполнить позже в Excel)")
            return None

        # Строгая проверка регулярным выражением на соответствие стандарту ISO
        if re.match(r'^\d{4}-\d{2}-\d{2}$', change_date):
            print(f"    ✓ Зафиксировано Change Date: '{change_date}'")
            return change_date

        print("    [Ошибка] Нарушен формат! Пожалуйста, введите дату строго как ГГГГ-ММ-ДД (например: 2026-09-14).")


def get_quantity_in_ss(
    df_matrix_current: pd.DataFrame,
    bp_number: str
) -> Dict[str, float]:
    """
    Интерактивный ввод количества деталей в Safety Stock для "старых" компонентов.
    Работает напрямую с двумерной матричной структурой df_matrix.

    Аргументы:
        df_matrix_current (pd.DataFrame): Текущая матрица Before/After.
        bp_number (str):                  Номер технического изменения (например, 'BP25010410').

    Возвращается:
        Dict[str, float]: Словарь соответствий {Part No. Before: количество в SS}.
    """
    print(f"\n  [Анализ деталей] Проверка остатков в Safety Stock для {bp_number}...")

    # Извлекаем чистые номера деталей Before через утилиту etl.py
    valid_parts = sorted(get_unique_non_empty_values(df_matrix_current['Part No. Before'], 'Part No. Before'))
    total_parts = len(valid_parts)

    if total_parts == 0:
        print("\n  [Информация] В структуре текущего BP не обнаружено старых деталей. Шаг автоматически пропускается.")
        return {}

    # Единый информационный блок: инструкция + список деталей
    print(f"\n  Ввод количества деталей в SS для BP {bp_number}")
    print("=" * 60)
    print("  ИНСТРУКЦИЯ:")
    print("    1. Введите номер детали из списка ниже (можно копировать).")
    print("    2. Укажите фактическое количество на складе (целое число ≥ 0).")
    print("    3. Нажмите Enter без ввода номера — всем оставшимся деталям")
    print("       автоматически присвоится 0 шт.")
    print("=" * 60)
    print("  СПИСОК СТАРЫХ ДЕТАЛЕЙ, ТРЕБУЮЩИХ УЧЕТА ОСТАТКОВ:")
    for i, part_no in enumerate(valid_parts, 1):
        print(f"    {i:>2}. {part_no}")
    print("=" * 60)

    quantity_dict: Dict[str, float] = {}
    remaining_parts = set(valid_parts)
    processed_count = 0

    # Запускаем цикл интерактивного опроса по оставшимся деталям
    while remaining_parts:
        print(f"\n  Прогресс: обработано {processed_count}/{total_parts} | осталось {len(remaining_parts)}")
        print("  Доступные номера деталей:")
        for i, part_no in enumerate(sorted(remaining_parts), 1):
            print(f"    {i}. {part_no}")

        # Используем безопасный системный ввод вместо сырого input()
        part_input = ask_user_input(
            "\n  Введите номер детали (или Enter — автозаполнение оставшихся нулями): "
        ).strip()

        # Если оператор нажал Enter на пустой строке — забиваем остаток списка нулями
        if not part_input:
            print("\n  [Автозаполнение] Для всех оставшихся деталей устанавливается количество: 0.0 шт.")
            for part_no in remaining_parts:
                quantity_dict[part_no] = 0.0
                print(f"    • {part_no} → 0.0 шт.")
            break

        # Защита от дурака: проверяем, присутствует ли введенный номер в наборе оставшихся
        if part_input not in remaining_parts:
            print(f"  [Ошибка] Номер '{part_input}' отсутствует в списке доступных старых деталей!")
            print("  Пожалуйста, скопируйте корректный номер из перечня выше.")
            continue

        # Внутренний цикл валидации числового значения
        while True:
            qty_input = ask_user_input(f"  Введите количество в SS для детали {part_input}: ").strip()

            if not qty_input:
                quantity = 0.0
                print("    → Принято значение по умолчанию: 0.0 шт.")
                break

            if qty_input.isdigit():
                quantity = int(qty_input)
                print(f"    ✓ Успешно зафиксировано: {quantity} шт.")
                break

            print("  [Ошибка] Некорректный ввод! Укажите целое положительное число или нажмите Enter.")

        quantity_dict[part_input] = quantity
        remaining_parts.remove(part_input)
        processed_count += 1

    # Выводим сводную верификационную таблицу для текущего шага
    total_ss = sum(quantity_dict.values())
    print("\n" + "=" * 60)
    print("  ИТОГОВЫЙ СВОД ОСТАТКОВ SAFETY STOCK ДЛЯ СЕССИИ:")
    print("=" * 60)
    for part_no in valid_parts:
        qty = quantity_dict.get(part_no, 0.0)
        print(f"    • {part_no}: {qty} шт.")
    print("-" * 60)
    print(f"  ВСЕГО ДЕТАЛЕЙ: {total_parts} | ОБЩЕЕ КОЛИЧЕСТВО В SS: {total_ss} шт.")
    print("=" * 60)

    return quantity_dict


def get_translation(
    series_list: List[pd.Series],
    field_name: str
) -> Dict[str, str]:
    """
    Интерактивный ввод переводов для уникальных текстовых значений.
    Самостоятельно извлекает уникальные непустые значения из переданных серий,
    объединяет их во множество и запускает пошаговый опрос оператора.

    Аргументы:
        series_list (List[pd.Series]): Список колонок (Series) для сбора уникальных значений.
                                       Может содержать одну или несколько серий
                                       (например, Before и After одновременно).
        field_name (str):              Название поля для вывода инструкций
                                       (например, "названий деталей").

    Возвращается:
        Dict[str, str]: Словарь соответствий {оригинал: перевод}.
    """
    # Собираем уникальные непустые значения из всех переданных серий
    all_unique_values: List[str] = []
    for series in series_list:
        values = get_unique_non_empty_values(series, field_name)
        all_unique_values.extend(values)

    # Убираем дубликаты между сериями и сортируем
    data = sorted(set(all_unique_values))

    if not data:
        print(f"\n  [Информация] В поле '{field_name}' отсутствуют данные для перевода. Шаг пропущен.")
        return {}

    translation_map: Dict[str, str] = {}

    print(f"\n  === Запуск интерактивного перевода: {field_name} ===")
    print(f"  Найдено уникальных записей к обработке: {len(data)}")
    print("  ИНСТРУКЦИЯ:")
    print("    • Введите перевод на русский язык и нажмите Enter.")
    print("    • Если оставить строку пустой (просто нажать Enter), значение останется оригинальным.")
    print("-" * 60)

    for i, value in enumerate(data, 1):
        # Защита от пустых ячеек и системных прочерков
        if pd.isna(value) or str(value).strip() in ('', '-'):
            translation_map[value] = '-'
            continue

        print(f"\n  [{i}/{len(data)}] Оригинал: {value}")

        user_input = ask_user_input("  Введите перевод: ").strip()

        if user_input == '':
            # Пользователь решил оставить оригинальный текст без изменений
            translation_map[value] = value
            print(f"    → Оставлен оригинал: '{value}'")
        else:
            # Чистая замена текста на русский перевод
            translation_map[value] = user_input
            print(f"    → Установлен перевод: '{user_input}'")

    return translation_map


def get_supplier_localization_status(
    series_list: List[pd.Series],
    bp_number: str
) -> Dict[str, str]:
    """
    Интерактивный ввод статуса локализации для каждого уникального поставщика сессии.
    Самостоятельно извлекает уникальные непустые значения из переданных серий
    и заполняет колонки Localization Before / After.

    Доступные статусы:
    - 1: Да (локальный поставщик)   -> "Да"
    - 2: Нет (зарубежный поставщик) -> "Нет"
    - Enter: статус не указан       -> "-"

    Аргументы:
        series_list (List[pd.Series]): Список колонок для сбора уникальных поставщиков
                                       (обычно [Supplier Name Before, Supplier Name After]).
        bp_number (str):               Номер технического изменения (например, 'BP25010410').

    Возвращается:
        Dict[str, str]: Словарь соответствий {Название поставщика: Статус локализации}.
    """
    print(f"\n  [Анализ логистики] Проверка статусов локализации контрагентов для {bp_number}...")

    # Собираем уникальные непустые значения поставщиков из всех переданных серий
    all_unique_values: List[str] = []
    for series in series_list:
        values = get_unique_non_empty_values(series, 'Supplier Name')
        all_unique_values.extend(values)

    # Убираем дубликаты между сериями и сортируем
    all_unique_suppliers = sorted(set(all_unique_values))
    total_suppliers = len(all_unique_suppliers)

    if total_suppliers == 0:
        print("  [Информация] В структуре текущего изменения не обнаружено заполненных поставщиков. Шаг пропущен.")
        return {}

    print(f"\n  Найдено уникальных заводов-поставщиков: {total_suppliers}")
    print("-" * 60)
    print("  СПИСОК ПОСТАВЩИКОВ ДЛЯ ОПРЕДЕЛЕНИЯ ЛОКАЛИЗАЦИИ:")
    for i, supplier in enumerate(all_unique_suppliers, 1):
        print(f"    {i}. {supplier}")
    print("-" * 60)
    print("  ДОСТУПНЫЕ КОДЫ СТАТУСОВ:")
    print("    1 - Да (Локальный поставщик)  -> 'Да'")
    print("    2 - Нет (Импортный поставщик) -> 'Нет'")
    print("    [Enter] - Оставить пустым     -> '-'")
    print("-" * 60)

    localization_map: Dict[str, str] = {}

    # Запускаем пошаговый опрос без KeyboardInterrupt блоков
    for i, supplier in enumerate(all_unique_suppliers, 1):
        print(f"\n  [{i}/{total_suppliers}] Контрагент: {supplier}")

        while True:
            choice = ask_user_input("  Введите код статуса (1, 2 или Enter): ").strip()

            if choice == '':
                localization_map[supplier] = '-'
                print("    → Статус оставлен неопределенным ('-')")
                break
            elif choice == '1':
                localization_map[supplier] = 'Да'
                print("    ✓ Зафиксировано: Да")
                break
            elif choice == '2':
                localization_map[supplier] = 'Нет'
                print("    ✓ Зафиксировано: Нет")
                break

            print("  [Ошибка] Неверный выбор! Пожалуйста, введите цифру 1, 2 или просто нажмите Enter.")

    return localization_map


def get_supplier_location_data(
    series_list: List[pd.Series],
    bp_number: str,
    branch: str
) -> Dict[str, Dict[str, str]]:
    """
    Интерактивный ввод данных о местоположении для каждого уникального поставщика.

    Самостоятельно извлекает уникальные непустые значения из переданных серий
    и последовательно запрашивает 4 поля: страна, город, улица, здание.

    Аргументы:
        series_list (List[pd.Series]): Список колонок для сбора уникальных поставщиков.
        bp_number (str):               Номер технического изменения (например, 'BP25010410').
        branch (str):                  Ветка матрицы ('Before' или 'After') для
                                       корректного вывода инструкций оператору.

    Возвращается:
        Dict[str, Dict[str, str]]: Словарь из 4 плоских мап вида
            {
                'location': {'Поставщик А': 'Китай',   ...},
                'city':     {'Поставщик А': 'Шанхай',  ...},
                'street':   {'Поставщик А': 'Nanjing', ...},
                'building': {'Поставщик А': '123',     ...},
            }
        Пустые поля заполняются значением '-'.
    """
    print(f"\n  [Анализ логистики] Ввод местоположения {branch}-поставщиков для {bp_number}...")

    # Собираем уникальные непустые значения поставщиков из всех переданных серий
    all_unique_values: List[str] = []
    for series in series_list:
        values = get_unique_non_empty_values(series, f'Supplier Name {branch}')
        all_unique_values.extend(values)

    # Убираем дубликаты между сериями и сортируем
    all_unique_suppliers = sorted(set(all_unique_values))
    total_suppliers = len(all_unique_suppliers)

    if total_suppliers == 0:
        print(f"  [Информация] В структуре текущего изменения не обнаружено заполненных "
              f"{branch}-поставщиков. Шаг пропущен.")
        return {}

    print(f"\n  Найдено уникальных {branch}-поставщиков: {total_suppliers}")
    print("-" * 60)
    print(f"  СПИСОК {branch.upper()}-ПОСТАВЩИКОВ ДЛЯ ВВОДА МЕСТОПОЛОЖЕНИЯ:")
    for i, supplier in enumerate(all_unique_suppliers, 1):
        print(f"    {i}. {supplier}")
    print("-" * 60)
    print("  ИНСТРУКЦИЯ:")
    print("    • Для каждого поставщика последовательно введите 4 поля:")
    print("        1. Страна           (Supplier Location)")
    print("        2. Город            (Supplier City)")
    print("        3. Улица            (Supplier Street)")
    print("        4. Здание           (Supplier Building)")
    print("    • Пустая строка (просто Enter) оставляет поле неопределённым ('-').")
    print("-" * 60)

    # 4 плоских мапы: {поставщик: значение}
    location_map: Dict[str, str] = {}
    city_map:     Dict[str, str] = {}
    street_map:   Dict[str, str] = {}
    building_map: Dict[str, str] = {}

    # Запускаем пошаговый опрос без KeyboardInterrupt блоков
    for i, supplier in enumerate(all_unique_suppliers, 1):
        print(f"\n  [{i}/{total_suppliers}] Контрагент: {supplier}")

        location = ask_user_input("    1. Страна           (Location): ").strip()
        city     = ask_user_input("    2. Город            (City):     ").strip()
        street   = ask_user_input("    3. Улица            (Street):   ").strip()
        building = ask_user_input("    4. Здание           (Building): ").strip()

        location_map[supplier] = location if location else '-'
        city_map[supplier]     = city     if city     else '-'
        street_map[supplier]   = street   if street   else '-'
        building_map[supplier] = building if building else '-'

        # Краткая сводка по введённым данным для визуального контроля
        print(f"    ✓ Зафиксировано: {location_map[supplier]} | "
              f"{city_map[supplier]} | {street_map[supplier]} | {building_map[supplier]}")

    return {
        'location': location_map,
        'city':     city_map,
        'street':   street_map,
        'building': building_map,
    }


def detect_workshop_by_code(
    workcenter_code: str
) -> Dict[str, str]:
    """
    Определяет код и название цеха по коду воркцентра (рабочей станции).

    Логика:
        - Код воркцентра имеет паттерн [Завод][Цех][Номер], например 'HAD112'.
        - Первые две буквы (например 'HA') однозначно определяют цех.
        - По префиксу находится код цеха через WORKSHOP_PREFIX_TO_CODE.
        - По коду цеха находится название через WORKSHOPS.

    Аргументы:
        workcenter_code (str): Код воркцентра, например 'HAD112'.

    Возвращается:
        Dict[str, str]: Словарь вида
            {'workshop_code': 'AS', 'workshop_name': 'Assembling'}.
            Если код не распознан — оба поля равны '-'.
    """
    empty_result = {'workshop_code': '-', 'workshop_name': '-'}

    if not isinstance(workcenter_code, str) or len(workcenter_code) < 2:
        return empty_result

    # Извлекаем префикс (первые две буквы)
    prefix = workcenter_code[:2].upper()

    workshop_code = WORKSHOP_PREFIX_TO_CODE.get(prefix)
    if not workshop_code:
        return empty_result

    workshop_name = WORKSHOPS.get(workshop_code, {}).get('name', '-')

    return {
        'workshop_code': workshop_code,
        'workshop_name': workshop_name,
    }


def load_packing_list_file(
    filename: str,
    packing_dir: str = INPUT_PACKING_LIST_DIR
) -> Optional[pd.DataFrame]:
    """
    Загружает упаковочный лист по имени файла из директории input_packing_list_files/.

    Аргументы:
        filename (str): Имя файла (с расширением или без).
        packing_dir (str): Директория с упаковочными листами.

    Возвращается:
        Optional[pd.DataFrame]: DataFrame упаковочного листа или None при ошибке.
    """
    if not filename or not filename.strip():
        return None

    filename = filename.strip()

    # Если расширение не указано — добавляем .xlsx
    if not filename.lower().endswith(('.xlsx', '.xls')):
        filename = filename + '.xlsx'

    full_path = os.path.join(packing_dir, filename)

    if not os.path.exists(full_path):
        print(f"  [Ошибка] Файл '{filename}' не найден в '{packing_dir}'.")
        return None

    try:
        df_batch = pd.read_excel(full_path, sheet_name=0)
        print(f"  [Упаковка] Файл '{filename}' загружен ({df_batch.shape[0]} строк).")
        return df_batch
    except PermissionError:
        print(f"  [Ошибка] Нет прав для чтения файла '{filename}'. Закройте его, если открыт в Excel.")
        return None
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
        print(f"  [Ошибка] Файл '{filename}' повреждён: {e}")
        return None
    except Exception as e:
        print(f"  [Ошибка] Не удалось прочитать файл '{filename}': {e}")
        return None


def parse_dimensions(
    size_str: str
) -> tuple:
    """
    Парсит строку размера в формате "Д×Ш×В" или "ДxШxВ" в кортеж из трёх float.

    Поддерживаемые разделители:
        - × (U+00D7, знак умножения)
        - x / X (латинская)

    Аргументы:
        size_str (str): Строка размера, например "550×380×230".

    Возвращается:
        tuple: (length, width, height) как float или (np.nan, np.nan, np.nan),
               если строку распарсить не удалось.
    """
    nan_result = (np.nan, np.nan, np.nan)

    if not isinstance(size_str, str) or not size_str.strip():
        return nan_result

    # Нормализуем разделители: заменяем × и X на x
    normalized = size_str.strip().replace('×', 'x').replace('X', 'x')

    # Разбиваем по 'x'
    parts = [p.strip() for p in normalized.split('x')]

    if len(parts) != 3:
        return nan_result

    try:
        length = float(parts[0])
        width = float(parts[1])
        height = float(parts[2])
        return (length, width, height)
    except (ValueError, TypeError):
        return nan_result


def collect_unique_parts(
    df: pd.DataFrame,
    part_col: str,
    name_col: str
) -> Tuple[Dict[str, List[Hashable]], Dict[str, str]]:
    """
    Собирает уникальные детали и их индексы из матрицы Before/After.

    Одним проходом формирует два словаря:
        - parts_map: {part_no: [индексы строк]}
        - names_map: {part_no: название детали}

    Аргументы:
        df (pd.DataFrame): Текущая матрица Before/After.
        part_col (str):    Имя колонки с номером детали
                           ('Part No. Before' / 'Part No. After').
        name_col (str):    Имя колонки с названием детали
                           ('Part Name Before' / 'Part Name After').

    Возвращается:
        Tuple[Dict[str, List[Hashable]], Dict[str, str]]:
            - parts_map: {part_no: [индексы]}
            - names_map: {part_no: имя}
    """
    parts_map: Dict[str, List[Hashable]] = {}
    names_map: Dict[str, str] = {}

    for idx, row in df.iterrows():
        part_no = str(row.get(part_col, '')).strip()
        if not part_no or part_no == '-':
            continue

        part_name = str(row.get(name_col, '')).strip()
        if part_no not in parts_map:
            parts_map[part_no] = []
            names_map[part_no] = part_name
        parts_map[part_no].append(idx)

    return parts_map, names_map


def calculate_packaging_metrics(
    packaging: Dict[str, Any]
) -> Dict[str, float]:
    """
    Рассчитывает производные упаковочные метрики (площади/объёмы коробки и паллеты).

    Аргументы:
        packaging (Dict[str, Any]): Словарь от extract_packaging_data() с ключами
                                    'box_size' и 'pallet_size'.

    Возвращается:
        Dict[str, float]: Словарь с ключами:
            - 'box_len', 'box_w', 'box_h'
            - 'box_space', 'box_volume'
            - 'pal_len', 'pal_w', 'pal_h'
            - 'pal_space', 'pal_volume'
        Все значения — float или np.nan.
    """
    box_size_str = str(packaging.get('box_size') or '')
    pallet_size_str = str(packaging.get('pallet_size') or '')

    box_len, box_w, box_h = parse_dimensions(box_size_str)
    pal_len, pal_w, pal_h = parse_dimensions(pallet_size_str)

    box_space = round(box_len * box_w / 1_000_000, 2) if not any(pd.isna([box_len, box_w])) else np.nan
    box_volume = round(box_len * box_w * box_h / 1_000_000_000, 2) if not any(pd.isna([box_len, box_w, box_h])) else np.nan
    pal_space = round(pal_len * pal_w / 1_000_000, 2) if not any(pd.isna([pal_len, pal_w])) else np.nan
    pal_volume = round(pal_len * pal_w * pal_h / 1_000_000_000, 2) if not any(pd.isna([pal_len, pal_w, pal_h])) else np.nan

    return {
        'box_len': box_len, 'box_w': box_w, 'box_h': box_h,
        'box_space': box_space, 'box_volume': box_volume,
        'pal_len': pal_len, 'pal_w': pal_w, 'pal_h': pal_h,
        'pal_space': pal_space, 'pal_volume': pal_volume,
    }


def get_packaging_user_inputs(
    part_no: str,
    part_name: str,
    branch: str
) -> Dict[str, Any]:
    """
    Интерактивный ввод пользовательских упаковочных полей для одной детали:
        - Box Type        (меню PACKAGING_TYPE)
        - Box Stacking    (int ≥ 1 или Enter → np.nan)
        - Pallet Type     (меню PACKAGING_TYPE)
        - Pallet Stacking (int ≥ 1 или Enter → np.nan)

    Аргументы:
        part_no (str):   Номер детали.
        part_name (str): Название детали.
        branch (str):    Ветка ('Before' / 'After').

    Возвращается:
        Dict[str, Any]: Словарь с ключами:
            - 'box_type' (str | float):      Тип коробки или np.nan
            - 'box_stacking' (int | float):  Штабелирование или np.nan
            - 'pallet_type' (str | float):   Тип паллеты или np.nan
            - 'pallet_stacking' (int | float): Штабелирование паллеты или np.nan
    """
    print(f"\n  [{branch}] Упаковочные параметры для {part_no}")
    display_name = part_name[:70] + '...' if len(part_name) > 70 else part_name
    print(f"  Название: {display_name}")

    # --- Box Type ---
    print(f"\n  Доступные типы упаковки: {', '.join(PACKAGING_TYPE)}")
    while True:
        box_type = ask_user_input(
            "  Box Type [Enter — пропустить]: "
        ).strip()
        if not box_type:
            box_type = np.nan
            print("    → Оставлено пустым (np.nan)")
            break
        if box_type.lower() in [t.lower() for t in PACKAGING_TYPE]:
            print(f"    ✓ Зафиксировано Box Type: '{box_type}'")
            break
        print(f"  [Ошибка] Недопустимый тип. Выберите из: {', '.join(PACKAGING_TYPE)}")

    # --- Box Stacking ---
    box_stacking = ask_int_or_nan("  Box Stacking (units, целое ≥ 1) [Enter — пропустить]: ")

    # --- Pallet Type ---
    print(f"\n  Доступные типы паллет: {', '.join(PACKAGING_TYPE)}")
    while True:
        pallet_type = ask_user_input(
            "  Pallet Type [Enter — пропустить]: "
        ).strip()
        if not pallet_type:
            pallet_type = np.nan
            print("    → Оставлено пустым (np.nan)")
            break
        if pallet_type.lower() in [t.lower() for t in PACKAGING_TYPE]:
            print(f"    ✓ Зафиксировано Pallet Type: '{pallet_type}'")
            break
        print(f"  [Ошибка] Недопустимый тип. Выберите из: {', '.join(PACKAGING_TYPE)}")

    # --- Pallet Stacking ---
    pallet_stacking = ask_int_or_nan("  Pallet Stacking (units, целое ≥ 1) [Enter — пропустить]: ")

    return {
        'box_type': box_type,
        'box_stacking': box_stacking,
        'pallet_type': pallet_type,
        'pallet_stacking': pallet_stacking,
    }


def apply_packaging_to_rows(
    df: pd.DataFrame,
    parts_map: Dict[str, List[Hashable]],
    part_no: str,
    packaging: Optional[Dict[str, Any]],
    user_inputs: Dict[str, Any],
    branch: str
) -> None:
    """
    Заполняет все 18 упаковочных колонок для одной детали и всех её дубликатов.

    Источники данных:
        - packaging (Dict | None):   данные из упаковочного листа + расчёты;
        - user_inputs (Dict):        пользовательский ввод (тип, штабелирование).

    Если packaging is None — все поля из файла заполняются np.nan.
    Если какое-то поле в packaging == np.nan — соответствующая колонка = np.nan.
    Box Net Weight считается как part_net_weight × parts_qty_box; если любой из
    сомножителей np.nan — результат np.nan.

    Аргументы:
        df (pd.DataFrame):          Матрица для заполнения.
        parts_map (Dict):           {part_no: [индексы строк]}.
        part_no (str):              Номер детали.
        packaging (Optional[Dict]): Словарь от extract_packaging_data() или None.
        user_inputs (Dict):         Словарь от get_packaging_user_inputs().
        branch (str):               Ветка ('Before' / 'After').
    """
    # Значения по умолчанию (если упаковочный лист не найден)
    parts_qty_box = np.nan
    boxes_per_pallet = np.nan
    pallet_gross_weight = np.nan
    part_net_weight = np.nan
    metrics: Dict[str, float] = {
        'box_len': np.nan, 'box_w': np.nan, 'box_h': np.nan,
        'box_space': np.nan, 'box_volume': np.nan,
        'pal_len': np.nan, 'pal_w': np.nan, 'pal_h': np.nan,
        'pal_space': np.nan, 'pal_volume': np.nan,
    }

    # Источник A: упаковочный лист + расчёт
    if packaging is not None:
        metrics = calculate_packaging_metrics(packaging)
        parts_qty_box = packaging.get('parts_qty_box', np.nan)
        boxes_per_pallet = packaging.get('boxes_per_pallet', np.nan)
        pallet_gross_weight = packaging.get('pallet_gross_weight', np.nan)
        part_net_weight = packaging.get('part_net_weight', np.nan)

    # Расчёт Box Net Weight = part_net_weight × parts_qty_box
    if pd.notna(part_net_weight) and pd.notna(parts_qty_box):
        box_net_weight = part_net_weight * parts_qty_box
    else:
        box_net_weight = np.nan

    # Заполняем все строки детали (включая дубликаты)
    for idx in parts_map[part_no]:
        df.at[idx, f'Quantity per Box {branch}'] = parts_qty_box
        df.at[idx, f'Box Type {branch}'] = user_inputs['box_type']
        df.at[idx, f'Box Length {branch} mm'] = metrics['box_len']
        df.at[idx, f'Box Width {branch} mm'] = metrics['box_w']
        df.at[idx, f'Box Height {branch} mm'] = metrics['box_h']
        df.at[idx, f'Box Area {branch} m2'] = metrics['box_space']
        df.at[idx, f'Box Volume {branch} m3'] = metrics['box_volume']
        df.at[idx, f'Box Net Weight {branch} kg'] = box_net_weight
        df.at[idx, f'Box Stacking {branch} units'] = user_inputs['box_stacking']
        df.at[idx, f'Boxes per Pallet {branch} units'] = boxes_per_pallet
        df.at[idx, f'Pallet Type {branch}'] = user_inputs['pallet_type']
        df.at[idx, f'Pallet Length {branch} mm'] = metrics['pal_len']
        df.at[idx, f'Pallet Width {branch} mm'] = metrics['pal_w']
        df.at[idx, f'Pallet Height {branch} mm'] = metrics['pal_h']
        df.at[idx, f'Pallet Area {branch} m2'] = metrics['pal_space']
        df.at[idx, f'Pallet Volume {branch} m3'] = metrics['pal_volume']
        df.at[idx, f'Pallet Gross Weight {branch} kg'] = pallet_gross_weight
        df.at[idx, f'Pallet Stacking {branch} units'] = user_inputs['pallet_stacking']


def copy_packaging_before_to_after(
    df: pd.DataFrame,
    before_parts: Dict[str, List[Hashable]],
    after_parts: Dict[str, List[Hashable]]
) -> Dict[str, str]:
    """
    Автоматически копирует все 18 упаковочных полей Before → After
    для After-деталей, чей номер совпадает с номером Before-детали.

    Копирование выполняется только для пар с одинаковым Part No.:
    это означает, что деталь не изменилась, и её упаковка идентична.

    Аргументы:
        df (pd.DataFrame):            Матрица Before/After.
        before_parts (Dict):          {part_no: [индексы Before-строк]}.
        after_parts (Dict):           {part_no: [индексы After-строк]}.

    Возвращается:
        Dict[str, str]: {part_no: part_no} — словарь After-деталей,
                        для которых выполнено автокопирование.
    """
    copied_parts: Dict[str, str] = {}

    for part_no, after_idx_list in after_parts.items():
        if part_no not in before_parts:
            continue

        # Берём упаковочные данные из первой Before-строки этой детали
        # (все строки одной детали имеют одинаковые упаковочные поля)
        before_idx = before_parts[part_no][0]

        # Копируем все 18 полей во все After-строки этой детали (включая дубликаты)
        for idx in after_idx_list:
            df.at[idx, 'Quantity per Box After'] = df.at[before_idx, 'Quantity per Box Before']
            df.at[idx, 'Box Type After'] = df.at[before_idx, 'Box Type Before']
            df.at[idx, 'Box Length After mm'] = df.at[before_idx, 'Box Length Before mm']
            df.at[idx, 'Box Width After mm'] = df.at[before_idx, 'Box Width Before mm']
            df.at[idx, 'Box Height After mm'] = df.at[before_idx, 'Box Height Before mm']
            df.at[idx, 'Box Area After m2'] = df.at[before_idx, 'Box Area Before m2']
            df.at[idx, 'Box Volume After m3'] = df.at[before_idx, 'Box Volume Before m3']
            df.at[idx, 'Box Net Weight After kg'] = df.at[before_idx, 'Box Net Weight Before kg']
            df.at[idx, 'Box Stacking After units'] = df.at[before_idx, 'Box Stacking Before units']
            df.at[idx, 'Boxes per Pallet After units'] = df.at[before_idx, 'Boxes per Pallet Before units']
            df.at[idx, 'Pallet Type After'] = df.at[before_idx, 'Pallet Type Before']
            df.at[idx, 'Pallet Length After mm'] = df.at[before_idx, 'Pallet Length Before mm']
            df.at[idx, 'Pallet Width After mm'] = df.at[before_idx, 'Pallet Width Before mm']
            df.at[idx, 'Pallet Height After mm'] = df.at[before_idx, 'Pallet Height Before mm']
            df.at[idx, 'Pallet Area After m2'] = df.at[before_idx, 'Pallet Area Before m2']
            df.at[idx, 'Pallet Volume After m3'] = df.at[before_idx, 'Pallet Volume Before m3']
            df.at[idx, 'Pallet Gross Weight After kg'] = df.at[before_idx, 'Pallet Gross Weight Before kg']
            df.at[idx, 'Pallet Stacking After units'] = df.at[before_idx, 'Pallet Stacking Before units']

        copied_parts[part_no] = part_no

    return copied_parts


@lru_cache(maxsize=1)
def load_configuration_file(
    config_dir: str = INPUT_CONFIGURATION_DIR
) -> Optional[pd.DataFrame]:
    """
    Ленивая загрузка конфигурационного файла configuration.xlsx из
    input_files/input_configuration_files/.

    Функция декорирована lru_cache(maxsize=1) — файл читается один раз
    за сессию, повторные вызовы возвращают закэшированный результат.
    Для сброса кэша используйте load_configuration_file.cache_clear().

    Возвращается:
        Optional[pd.DataFrame]: DataFrame с конфигурацией или None, если файл
                                не найден или не удалось прочитать.
    """
    latest_file = find_latest_excel_file(
        file_prefix=CONFIGURATION_PREFIX,
        file_path=config_dir
    )

    if not latest_file:
        print(f"  [Конфигурация] Файл не найден в '{config_dir}'.")
        return None

    required_cols = CONF_TO_KEEP_COLS

    df_raw = read_excel_file(
        filename=latest_file,
        file_path=config_dir,
        sheet_name='common',
        cols=required_cols
    )

    if df_raw is None or df_raw.empty:
        print("  [Конфигурация] Не удалось прочитать данные или файл пуст.")
        return None

    df_clean = normalize_data(df_raw)
    print(f"  [Конфигурация] Загружено {len(df_clean)} записей из '{latest_file}'.")
    return df_clean


def process_bp_file(
    bp_filename: str
) -> Optional[Dict[str, Any]]:
    """
    Конвейер сквозной интерактивной обработки одного файла технического изменения (BP).

    Каждый шаг включает:
    - Интерактивное взаимодействие с пользователем
    - Возможность отката (retry)
    - Сохранение состояния

    Аргументы:
        bp_filename (str): Имя файла BP для обработки.

    Возвращается:
        dict: Словарь вида {'bp_number': str, 'dataframe': pd.DataFrame}
              или None при критической ошибке.

    Шаги обработки:
        ШАГ 1: ЗАГРУЗКА, АУДИТ И НОРМАЛИЗАЦИЯ ДАННЫХ (BP<номер>.xlsx)
        ШАГ 2: РАЗДЕЛЕНИЕ ДЕТАЛЕЙ НА СТАРЫЕ И НОВЫЕ (BEFORE / AFTER)
        ШАГ 3: ОПРЕДЕЛЕНИЕ СТАТУСА (Status)
        ШАГ 4: ОПРЕДЕЛЕНИЕ НАЗВАНИЙ МОДЕЛЕЙ ПО КОДАМ (BOM Product Code -> BOM Product Name)
        ШАГ 5: ВВОД ФАКТИЧЕСКОЙ ПАРТИИ (Batch Fact)
        ШАГ 6: ВВОД ФАКТИЧЕСКОЙ ДАТЫ (Change Date)
        ШАГ 7: ВВОД КОЛИЧЕСТВА BEFORE-ДЕТАЛЕЙ В SAFETY STOCK (Quantity Parts in SS)
        ШАГ 8: ПЕРЕВОД НАЗВАНИЙ ДЕТАЛЕЙ (Part Name Before / After)
        ШАГ 9: ПЕРЕВОД ОФИЦИАЛЬНЫХ НАЗВАНИЙ ПОСТАВЩИКОВ (Supplier Name Before / After)
        ШАГ 10: ВВОД СТАТУСА ЛОКАЛИЗАЦИИ ПОСТАВЩИКОВ (Localization Before / After)
        ШАГ 11: ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ BEFORE-ПОСТАВЩИКОВ
        ШАГ 12: ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ AFTER-ПОСТАВЩИКОВ
        ШАГ 13: ПЕРЕВОД ОПИСАНИЯ К ИЗМЕНЕНИЮ (Change Description)
        ШАГ 14: ПЕРЕВОД РЕШЕНИЯ К ИЗМЕНЕНИЮ (Change Solution)
        ШАГ 15: ПЕРЕВОД ЦВЕТА (Color Code и Color Name)
        ШАГ 16: ОБРАБОТКА ЦЕХОВ И РАБОЧИХ СТАНЦИЙ (Workshop/Workcenter Code/Name Before/After)
        ШАГ 17: ПЕРЕВОД РЕБОВАНИЙ ПО УТИЛИЗАЦИИ СТАРЫХ ДЕТАЛЕЙ (Production Part Disposal)
        ШАГ 18: ПЕРЕВОД ТРЕБОВАНИЙ ПО ВЗАИМОЗАМЕНЯЕМОСТИ (Interchangeable)
        ШАГ 19: ЗАГРУЗКА КОНФИГУРАЦИОННОГО ФАЙЛА (YYYY-mm-dd_configuration.xlsx)
        ШАГ 20: РАСЧЁТ ПАРТИЙ В SAFETY STOCK (Quantity Batches in SS)
        ШАГ 21: ПОИСК КОНФИГУРАЦИИ ДЛЯ УТИЛИЗАЦИИ BEFORE-ДЕТАЛЕЙ
        ШАГ 22: ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ BEFORE-ДЕТАЛЕЙ
        ШАГ 23: ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ AFTER-ДЕТАЛЕЙ
        ШАГ 24: УПОРЯДОЧИВАНИЕ КОЛОНОК
        ШАГ 25: СОХРАНЕНИЕ РЕЗУЛЬТАТА
    """
    bp_start_time = time.time()
    bp_number = os.path.basename(bp_filename).replace('.xlsx', '').replace('.XLSX', '')

    bp_dir = os.path.dirname(bp_filename)
    bp_pure_name = os.path.basename(bp_filename)

    # Инициализируем сквозные переменные конвейера
    saved_state: Optional[pd.DataFrame] = None
    df_matrix: Optional[pd.DataFrame] = None

    # Счётчики для нумерации шагов
    bp_total_steps = 25
    counter = 0


    #=============== ШАГ 1: ЗАГРУЗКА, АУДИТ И НОРМАЛИЗАЦИЯ ДАННЫХ (BP<номер>.xlsx) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_pure_name}] ЗАГРУЗКА, АУДИТ И НОРМАЛИЗАЦИЯ ДАННЫХ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_raw: Optional[pd.DataFrame] = read_excel_file(filename=bp_pure_name, file_path=bp_dir)

        if df_raw is None or df_raw.empty:
            print(f"  [Критический сбой] Не удалось прочитать данные из файла {bp_pure_name}.")
            return None

        # Принудительно очищаем сами имена колонок от случайных пробелов оператора
        df_raw.columns = [str(col).strip() for col in df_raw.columns]

        available_cols = [col for col in BP_GMOM_TO_KEEP_COLS if col in df_raw.columns]
        missing_cols = set(BP_GMOM_TO_KEEP_COLS) - set(available_cols)

        if missing_cols:
            print("\n  " + "!" * 60)
            print(f"  [ВНИМАНИЕ] В структуре файла {bp_pure_name} не найдены обязательные колонки!")
            print("  " + "!" * 60)
            print("  Отсутствуют столбцы:")
            for col in missing_cols:
                print(f"    - {col}")

            print("\n  Варианты решения проблемы:")
            print("    Нажмите: 1 - Исправить заголовки вручную (откройте файл в Excel, переименуйте колонки по списку ниже и повторите проверку)")
            print("    Нажмите: 2 - Пропустить этот BP файл (конвейер перейдет к обработке следующего файла изменений)")

            print("\n  Обязательный целевой формат заголовков:")
            print(f"    {', '.join(STEP_FOCUS_COLUMNS_MAP[counter - 1])}")
            print("-" * 60)

            while True:
                user_choice = ask_user_input("  Выберите вариант действия (1 или 2): ").strip()
                if user_choice in ('1', '2'):
                    break
                print("  Некорректный ввод. Пожалуйста, введите цифру 1 или 2.")

            if user_choice == '1':
                print("\n  [Ожидание] Скрипт приостановлен.")
                print(f"  Пожалуйста, откройте файл '{os.path.join(bp_dir, bp_pure_name)}', скорректируйте имена колонок и сохраните его.")
                wait_for_user("  После сохранения файла нажмите Enter, чтобы запустить повторную загрузку...")
                continue
            elif user_choice == '2':
                print(f"\n  [Пропуск] Файл '{bp_pure_name}' исключен из текущей сессии обработки.")
                return {"status": "SKIP_FILE"}

        df_filtered: pd.DataFrame = df_raw[available_cols].copy()

        df_step = normalize_data(
            df_filtered,
            int_columns=BP_GBOM_INT_COLS,
            float_columns=None,
            datetime_columns=BP_GBOM_DATETIME_COLS
        )

        df_step['BP_No'] = bp_number

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break


    #=============== ШАГ 2: РАЗДЕЛЕНИЕ ДЕТАЛЕЙ НА СТАРЫЕ И НОВЫЕ (BEFORE / AFTER) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] РАЗДЕЛЕНИЕ ДЕТАЛЕЙ НА СТАРЫЕ И НОВЫЕ (BEFORE / AFTER)"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        list_of_pair_dicts = find_pairs(df_matrix)
        df_step = pd.DataFrame(list_of_pair_dicts)

        if df_step.empty:
            print(f"  [Критическая ошибка] Алгоритм find_pairs() вернул пустой массив для {bp_number}!")
            return None

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл начнется сначала, позволяя разделить детали на старые и новые заново


    #=============== ШАГ 3: ОПРЕДЕЛЕНИЕ СТАТУСА (Status) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ОПРЕДЕЛЕНИЕ СТАТУСА"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()
        status_value = get_bp_status(bp_number)
        df_step['Status'] = status_value

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл начнется сначала, позволяя указать Status заново


    #=============== ШАГ 4: ОПРЕДЕЛЕНИЕ НАЗВАНИЙ МОДЕЛЕЙ ПО КОДАМ (BOM Product Code -> BOM Product Name) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ОПРЕДЕЛЕНИЕ НАЗВАНИЙ МОДЕЛЕЙ ПО КОДАМ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # Собираем уникальные BOM Product Code из матрицы
        unique_bom_codes = sorted(get_unique_non_empty_values(df_step['BOM Product Code'], 'BOM Product Code'))

        # Строим словарь соответствий {BOM Product Code: BOM Product Name}
        code_to_name: Dict[str, str] = {}
        not_found_codes: set = set()

        for bom_code in unique_bom_codes:
            # Префикс — первые 3 символа (например, 'A01G-R-BOM' → 'A01')
            prefix = bom_code[:3]

            model_name = BOM_PRODUCT_MAP.get(prefix)
            if model_name:
                code_to_name[bom_code] = model_name
            else:
                code_to_name[bom_code] = '-'
                not_found_codes.add(bom_code)

        # Применяем словарь через .map() — быстро и без дублирования lookup'ов
        df_step['BOM Product Name'] = (
            df_step['BOM Product Code']
            .astype(str).str.strip()
            .map(code_to_name)
            .fillna('-')
        )

        # Единый дедуплицированный отчёт по ненайденным моделям
        if not_found_codes:
            print("\n  [Внимание] Не удалось определить модель для следующих BOM Product Code:")
            for code in sorted(not_found_codes):
                print(f"    - {code}")

        # Сводка
        found_count = sum(1 for name in df_step['BOM Product Name'] if name != '-')
        print(f"\n  Всего строк: {len(df_step)} | Определено моделей: {found_count} | Не найдено: {len(df_step) - found_count}")

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя повторить определение моделей


    #=============== ШАГ 5: ВВОД ФАКТИЧЕСКОЙ ПАРТИИ (Batch Fact) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ВВОД ФАКТИЧЕСКОЙ ПАРТИИ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()
        batch_fact_val = get_batch_fact(bp_number)
        df_step['Batch Fact'] = batch_fact_val

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} зафиксирован за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл начнется сначала, позволяя скорректировать Batch fact


    #=============== ШАГ 6: ВВОД ФАКТИЧЕСКОЙ ДАТЫ (Change Date) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ВВОД ФАКТИЧЕСКОЙ ДАТЫ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()
        change_date_val: Optional[str] = get_change_date(bp_number)
        df_step['Change Date'] = change_date_val if change_date_val else '-'

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл начнется заново, позволяя скорректировать Change Date


    #=============== ШАГ 7: ВВОД КОЛИЧЕСТВА BEFORE-ДЕТАЛЕЙ В SAFETY STOCK (Quantity Parts in SS) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ВВОД КОЛИЧЕСТВА BEFORE-ДЕТАЛЕЙ В SAFETY STOCK"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()
        quantity_ss_map: Dict[str, float] = get_quantity_in_ss(df_step, bp_number)
        df_step['Quantity Parts in SS'] = df_step['Part No. Before'].map(quantity_ss_map).fillna(0).astype(int)

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя переписать объемы Safety Stock с нуля


    #=============== ШАГ 8: ПЕРЕВОД НАЗВАНИЙ ДЕТАЛЕЙ (Part Name Before / After) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПЕРЕВОД НАЗВАНИЙ ДЕТАЛЕЙ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()
        parts_translation_map = get_translation(
            series_list=[
                df_step['Part Name Before'],
                df_step['Part Name After']
            ],
            field_name="названий деталей (Part Name Before/After)"
        )

        if parts_translation_map:
            df_step['Part Name Before'] = df_step['Part Name Before'].map(parts_translation_map).fillna(df_step['Part Name Before'])
            df_step['Part Name After'] = df_step['Part Name After'].map(parts_translation_map).fillna(df_step['Part Name After'])

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя скорректировать переводы деталей


    #=============== ШАГ 9: ПЕРЕВОД ОФИЦИАЛЬНЫХ НАЗВАНИЙ ПОСТАВЩИКОВ (Supplier Name Before / After) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПЕРЕВОД ОФИЦИАЛЬНЫХ НАЗВАНИЙ ПОСТАВЩИКОВ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # Передаём обе ветки матрицы — функция сама соберёт уникальные значения поставщиков
        suppliers_translation_map = get_translation(
            series_list=[
                df_step['Supplier Name Before'],
                df_step['Supplier Name After']
            ],
            field_name="названий поставщиков (Supplier Name Before/After)"
        )

        # Применяем полученную карту переводов одновременно к обеим колонкам матрицы.
        # Если перевода нет, сохраняем исходное название во избежание порчи данных.
        if suppliers_translation_map:
            df_step['Supplier Name Before'] = df_step['Supplier Name Before'].map(suppliers_translation_map).fillna(df_step['Supplier Name Before'])
            df_step['Supplier Name After'] = df_step['Supplier Name After'].map(suppliers_translation_map).fillna(df_step['Supplier Name After'])

        # Выводим сопоставление номеров деталей и их отрефакторенных поставщиков
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        # Фиксируем слепок оперативной памяти перед верификацией
        saved_state = save_state(df_step)

        # Интерактивный запрос на фиксацию изменений или сброс (retry)
        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            # Шаг успешно пройден, фиксируем состояние матрицы и выходим из цикла
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя скорректировать переводы поставщиков


    #=============== ШАГ 10: ВВОД СТАТУСА ЛОКАЛИЗАЦИИ ПОСТАВЩИКОВ (Localization Before / After) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ВВОД СТАТУСА ЛОКАЛИЗАЦИИ ПОСТАВЩИКОВ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        localization_map = get_supplier_localization_status(
            series_list=[
                df_step['Supplier Name Before'],
                df_step['Supplier Name After']
            ],
            bp_number=bp_number
        )

        # Синхронно мапим результаты на обе ветки матрицы
        df_step['Localization Before'] = df_step['Supplier Name Before'].map(localization_map).fillna('-')
        df_step['Localization After'] = df_step['Supplier Name After'].map(localization_map).fillna('-')

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя переписать статусы локализации с нуля


    #=============== ШАГ 11: ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ BEFORE-ПОСТАВЩИКОВ ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ BEFORE-ПОСТАВЩИКОВ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        location_maps = get_supplier_location_data(
            series_list=[df_step['Supplier Name Before']],
            bp_number=bp_number,
            branch='Before'
        )

        # Нормализуем колонку поставщика один раз — убираем NaN/None из ключей
        supplier_keys = (
            df_step['Supplier Name Before']
            .astype(str)
            .str.strip()
        )

        # Мапим 4 плоские мапы на 4 целевые колонки Before-ветки
        df_step['Supplier Location Before'] = (
            supplier_keys.map(location_maps.get('location', {})).fillna('-')
        )
        df_step['Supplier City Before'] = (
            supplier_keys.map(location_maps.get('city', {})).fillna('-')
        )
        df_step['Supplier Street Before'] = (
            supplier_keys.map(location_maps.get('street', {})).fillna('-')
        )
        df_step['Supplier Building Before'] = (
            supplier_keys.map(location_maps.get('building', {})).fillna('-')
        )

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя переписать местоположение с нуля


    #=============== ШАГ 12: ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ AFTER-ПОСТАВЩИКОВ ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ AFTER-ПОСТАВЩИКОВ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        location_maps = get_supplier_location_data(
            series_list=[df_step['Supplier Name After']],
            bp_number=bp_number,
            branch='After'
        )

        # Нормализуем колонку поставщика один раз — убираем NaN/None из ключей
        supplier_keys = (
            df_step['Supplier Name After']
            .astype(str)
            .str.strip()
        )

        # Мапим 4 плоские мапы на 4 целевые колонки After-ветки
        df_step['Supplier Location After'] = (
            supplier_keys.map(location_maps.get('location', {})).fillna('-')
        )
        df_step['Supplier City After'] = (
            supplier_keys.map(location_maps.get('city', {})).fillna('-')
        )
        df_step['Supplier Street After'] = (
            supplier_keys.map(location_maps.get('street', {})).fillna('-')
        )
        df_step['Supplier Building After'] = (
            supplier_keys.map(location_maps.get('building', {})).fillna('-')
        )

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя переписать местоположение с нуля


    #=============== ШАГ 13: ПЕРЕВОД ОПИСАНИЯ К ИЗМЕНЕНИЮ (Change Description) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПЕРЕВОД ОПИСАНИЯ К ИЗМЕНЕНИЮ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # 1. Фильтрация: оставляем только китайский оригинал (источник истины)
        if 'Change Description' in df_step.columns:
            df_step['Change Description'] = df_step['Change Description'].apply(keep_chinese_lines)

            # 2. Интерактивный перевод очищенных значений
            desc_translation_map = get_translation(
                series_list=[df_step['Change Description']],
                field_name="описаний технического изменения (Change Description)"
            )

            # 3. Переносим перевод в целевую колонку и удаляем исходную
            if desc_translation_map:
                df_step['Change Description (RUS)'] = df_step['Change Description'].map(desc_translation_map).fillna('-')
            else:
                df_step['Change Description (RUS)'] = '-'

            df_step = df_step.drop(['Change Description'], axis=1)
        else:
            print("  [Информация] Колонка 'Change Description' отсутствует в матрице. Шаг пропущен.")
            df_step['Change Description (RUS)'] = '-'

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя скорректировать перевод описания

    #=============== ШАГ 14: ПЕРЕВОД РЕШЕНИЯ К ИЗМЕНЕНИЮ (Change Solution) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПЕРЕВОД РЕШЕНИЯ К ИЗМЕНЕНИЮ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # 1. Фильтрация: оставляем только китайский оригинал (источник истины)
        if 'Solution' in df_step.columns:
            df_step['Solution'] = df_step['Solution'].apply(keep_chinese_lines)

            # 2. Интерактивный перевод очищенных значений
            solution_translation_map = get_translation(
                series_list=[df_step['Solution']],
                field_name="решений технического изменения (Solution)"
            )

            # 3. Переносим перевод в целевую колонку 'Change Solution (RUS)' и удаляем исходную
            if solution_translation_map:
                df_step['Change Solution (RUS)'] = df_step['Solution'].map(solution_translation_map).fillna('-')
            else:
                df_step['Change Solution (RUS)'] = '-'

            df_step = df_step.drop(['Solution'], axis=1)
        else:
            print("  [Информация] Колонка 'Solution' отсутствует в матрице. Шаг пропущен.")
            df_step['Change Solution (RUS)'] = '-'

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя скорректировать перевод решения


    #=============== ШАГ 15: ПЕРЕВОД ЦВЕТА (Color Code и Color Name) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПЕРЕВОД ЦВЕТА"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # 1. Перевод названия цвета (Color Name → Color Name (RUS))
        if 'Color Name' in df_step.columns:
            color_translation_map = get_translation(
                series_list=[df_step['Color Name']],
                field_name="названий цвета (Color Name)"
            )

            if color_translation_map:
                df_step['Color Name (RUS)'] = df_step['Color Name'].map(color_translation_map).fillna('-')
            else:
                df_step['Color Name (RUS)'] = '-'

            df_step = df_step.drop(['Color Name'], axis=1)
        else:
            print("  [Информация] Колонка 'Color Name' отсутствует в матрице. Создаём 'Color Name (RUS)' со значениями '-'.")
            df_step['Color Name (RUS)'] = '-'

        # 2. Валидация колонки Color Code (не переводится, только проверка наличия)
        if 'Color Code' in df_step.columns:
            unique_codes = get_unique_non_empty_values(df_step['Color Code'], 'Color Code')

            if len(unique_codes) > 0:
                print(f"  [Color Code] Колонка содержит данные ({len(unique_codes)} уникальных кодов). Перевод не требуется.")
            else:
                print("  [Color Code] Колонка не содержит данных или все значения равны '-'. Заполняем '-'.")
                df_step['Color Code'] = '-'
        else:
            print("  [Color Code] Колонка отсутствует в матрице. Создаём 'Color Code' со значениями '-'.")
            df_step['Color Code'] = '-'

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя скорректировать перевод цвета


    #=============== ШАГ 16: ОБРАБОТКА ЦЕХОВ И РАБОЧИХ СТАНЦИЙ (Workshop/Workcenter Code/Name Before/After) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ОБРАБОТКА ЦЕХОВ И РАБОЧИХ СТАНЦИЙ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # Локальный set для дедупликации предупреждений о нераспознанных кодах
        unrecognized_codes: set = set()

        # Для обеих веток (Before / After) выполняем единый алгоритм:
        #   1. Извлекаем название воркцентра из скобок (если есть).
        #   2. По коду воркцентра определяем код и название цеха.
        for branch in ('Before', 'After'):
            col_wc_code = f'Workcenter Code {branch}'
            col_wc_name = f'Workcenter Name {branch}'

            # --- 1. Обработка названия воркцентра: извлекаем содержимое скобок ---
            if col_wc_name in df_step.columns:
                df_step[col_wc_name] = df_step[col_wc_name].apply(extract_parentheses_content)
            else:
                print(f"  [Информация] Колонка '{col_wc_name}' отсутствует в матрице. Создаём со значениями '-'.")
                df_step[col_wc_name] = '-'

            # --- 2. Определение цеха по коду воркцентра ---
            if col_wc_code in df_step.columns:
                workshop_codes: List[str] = []
                workshop_names: List[str] = []
                for code in df_step[col_wc_code]:
                    result = detect_workshop_by_code(code)

                    # Собираем нераспознанные валидные коды для единого отчёта
                    if result['workshop_code'] == '-' and isinstance(code, str) and code.strip() not in ('', '-'):
                        unrecognized_codes.add(code)

                    workshop_codes.append(result['workshop_code'])
                    workshop_names.append(result['workshop_name'])

                df_step[f'Workshop Code {branch}'] = workshop_codes
                df_step[f'Workshop Name {branch}'] = workshop_names
            else:
                print(f"  [Информация] Колонка '{col_wc_code}' отсутствует в матрице. Цех не определён.")
                df_step[f'Workshop Code {branch}'] = '-'
                df_step[f'Workshop Name {branch}'] = '-'

        # Единый дедуплицированный отчёт по нераспознанным кодам
        if unrecognized_codes:
            print("\n  [Внимание] Не удалось определить цех для следующих кодов воркцентров:")
            for code in sorted(unrecognized_codes):
                print(f"    - {code}")

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя скорректировать обработку цехов


    #=============== ШАГ 17: ПЕРЕВОД ТРЕБОВАНИЙ ПО УТИЛИЗАЦИИ СТАРЫХ ДЕТАЛЕЙ (Production Part Disposal) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПЕРЕВОД ТРЕБОВАНИЙ ПО УТИЛИЗАЦИИ СТАРЫХ ДЕТАЛЕЙ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # Перевод требований по утилизации (замена на месте, без сохранения оригинала)
        if 'Production Part Disposal' in df_step.columns:
            disposal_translation_map = get_translation(
                series_list=[df_step['Production Part Disposal']],
                field_name="требований по утилизации старых деталей (Production Part Disposal)"
            )

            if disposal_translation_map:
                df_step['Production Part Disposal'] = df_step['Production Part Disposal'].map(disposal_translation_map).fillna('-')
            else:
                df_step['Production Part Disposal'] = '-'
        else:
            print("  [Информация] Колонка 'Production Part Disposal' отсутствует в матрице. Создаём со значениями '-'.")
            df_step['Production Part Disposal'] = '-'

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя скорректировать перевод


    #=============== ШАГ 18: ПЕРЕВОД ТРЕБОВАНИЙ ПО ВЗАИМОЗАМЕНЯЕМОСТИ (Interchangeable) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПЕРЕВОД ТРЕБОВАНИЙ ПО ВЗАИМОЗАМЕНЯЕМОСТИ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # Перевод требований по взаимозаменяемости (замена на месте, без сохранения оригинала)
        if 'Interchangeable' in df_step.columns:
            interchangeable_translation_map = get_translation(
                series_list=[df_step['Interchangeable']],
                field_name="требований по взаимозаменяемости (Interchangeable)"
            )

            if interchangeable_translation_map:
                df_step['Interchangeable'] = df_step['Interchangeable'].map(interchangeable_translation_map).fillna('-')
            else:
                df_step['Interchangeable'] = '-'
        else:
            print("  [Информация] Колонка 'Interchangeable' отсутствует в матрице. Создаём со значениями '-'.")
            df_step['Interchangeable'] = '-'

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя скорректировать перевод


    #=============== ШАГ 19: ЗАГРУЗКА КОНФИГУРАЦИОННОГО ФАЙЛА (YYYY-mm-dd_configuration.xlsx) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ЗАГРУЗКА КОНФИГУРАЦИОННОГО ФАЙЛА"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_config = load_configuration_file()

        if df_config is None or df_config.empty:
            print(f"  [Внимание] Конфигурационный файл недоступен. Шаги {counter + 1} и {counter + 2} будут пропущены.")
            df_config = pd.DataFrame()

        # Показываем оператору превью загруженной конфигурации
        show_dataframe_preview(
            df_config,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        continue_flag, _, _ = confirm_step(
            step_name=step_name,
            df_current=df_config,
            saved_state=df_config
        )

        if continue_flag:
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        else:
            # При retry — сбрасываем кэш, чтобы перечитать файл заново
            print("  [Сброс кэша конфигурации] Файл будет перечитан при следующем входе в шаг.")
            load_configuration_file.cache_clear()
            continue


    #=============== ШАГ 20: РАСЧЁТ ПАРТИЙ В SAFETY STOCK (Quantity Batches in SS) ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] РАСЧЁТ ПАРТИЙ В SAFETY STOCK"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # Загружаем конфигурацию (из кэша, мгновенно)
        df_config = load_configuration_file()

        if df_config is None or df_config.empty:
            print("  [Пропуск] Конфигурация недоступна. Quantity Batches in SS = нет данных.")
            df_step['Quantity Batches in SS'] = np.nan
        else:
            # Строим словарь {BOM Product: Quantity Vehicle in Batch} один раз
            qty_map: Dict[str, float] = {}
            for _, row in df_config.iterrows():
                bom = str(row.get('BOM Product', '')).strip()
                if not bom or bom == '-':
                    continue
                try:
                    qty_map[bom] = float(row.get('Quantity Vehicle in Batch', 0))
                except (ValueError, TypeError):
                    qty_map[bom] = 0.0

            # Собираем BOM Product, для которых не нашлось значение
            not_found_products = {
                bom for bom in df_step['BOM Product'].astype(str).str.strip()
                if bom and bom != '-' and bom not in qty_map
            }

            if not_found_products:
                print("\n  [Внимание] Не найдено Quantity Vehicle in Batch для следующих BOM Product:")
                for product in sorted(not_found_products):
                    print(f"    - {product}")

            # Векторизованный расчёт Quantity Batches in SS
            qty_vehicle_series = (
                df_step['BOM Product'].astype(str).str.strip()
                .map(qty_map)
            )
            qty_in_ss_series = pd.to_numeric(df_step['Quantity Parts in SS'], errors='coerce')

            # Инициализируем колонку NaN
            df_step['Quantity Batches in SS'] = np.nan

            # Считаем только там, где qty_vehicle > 0 и qty_in_ss есть
            mask = (
                qty_vehicle_series.notna() & (qty_vehicle_series > 0) &
                qty_in_ss_series.notna()
            )
            df_step.loc[mask, 'Quantity Batches in SS'] = (
                qty_in_ss_series[mask] / qty_vehicle_series[mask]
            ).round(2)

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя пересчитать количество партий в Safety Stock


    #=============== ШАГ 21: ПОИСК КОНФИГУРАЦИИ ДЛЯ УТИЛИЗАЦИИ BEFORE-ДЕТАЛЕЙ ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПОИСК КОНФИГУРАЦИИ ДЛЯ УТИЛИЗАЦИИ BEFORE-ДЕТАЛЕЙ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # Загружаем конфигурацию (из кэша, мгновенно)
        df_config = load_configuration_file()

        if df_config is None or df_config.empty:
            print("  [Пропуск] Конфигурация недоступна. Заполняем целевые колонки значением '-'.")
            df_step['Configuration for Old Parts Using Out'] = '-'
            df_step['Batches for Old Parts Using Out'] = '-'
            df_step['Transmission'] = '-'
        else:
            # Инициализируем целевые колонки
            df_step['Configuration for Old Parts Using Out'] = '-'
            df_step['Batches for Old Parts Using Out'] = '-'
            df_step['Transmission'] = '-'

            # Достаём Batch Plan (одинаковый для всего BP)
            batch_plan = ''
            if not df_step.empty:
                batch_plan = str(df_step.iloc[0].get('Batch Plan', '')).strip()

            batch_plan_missing = (not batch_plan or batch_plan == '-' or batch_plan == 'nan')

            if batch_plan_missing:
                print("\n  " + "-" * 60)
                print("  ВНИМАНИЕ: BATCH PLAN НЕ ВВЕДЁН!")
                print("  Поиск конфигурации невозможен для полей:")
                print("    • Configuration for Old Parts Using Out")
                print("    • Batches for Old Parts Using Out")
                print("    • Transmission")
                print("  Эти поля останутся пустыми для ВСЕХ деталей в этом BP.")
                print("  " + "-" * 60)
            else:
                # Batch Plan[:3] — префикс для поиска в Batch Code
                batch_prefix = batch_plan[:3] if len(batch_plan) >= 3 else batch_plan

                # Строим индекс конфигурации: {BOM Product: DataFrame с совпадениями}
                # (для производительности при многих строках)

                for idx, row in df_step.iterrows():
                    bom_product = str(row.get('BOM Product', '')).strip()
                    if not bom_product or bom_product == '-':
                        continue

                    # Фильтруем по BOM Product
                    config_matches = df_config[
                        df_config['BOM Product'].astype(str).str.strip() == bom_product
                    ]
                    if config_matches.empty:
                        continue

                    # Фильтруем по Batch Code (startswith префикса)
                    config_match = config_matches[
                        config_matches['Batch Code'].astype(str).str.startswith(batch_prefix, na=False)
                    ]

                    if config_match.empty:
                        continue

                    # Configuration: берём ПЕРВОЕ значение
                    config_value = str(config_match.iloc[0].get('Configuration', '')).strip()
                    if config_value and config_value != 'nan':
                        df_step.at[idx, 'Configuration for Old Parts Using Out'] = config_value

                    # Batches и Transmission: собираем ВСЕ значения через '\n'
                    all_batch_codes = []
                    all_transmissions = []
                    for _, row_match in config_match.iterrows():
                        bc = str(row_match.get('Batch Code', '')).strip()
                        trans = str(row_match.get('Transmission', '')).strip()

                        # Пропускаем пустые
                        if (not bc or bc == 'nan') and (not trans or trans == 'nan'):
                            continue

                        if bc and bc != 'nan':
                            all_batch_codes.append(bc)
                        if trans and trans != 'nan':
                            all_transmissions.append(trans)

                    if all_batch_codes:
                        df_step.at[idx, 'Batches for Old Parts Using Out'] = '\n'.join(all_batch_codes)
                    if all_transmissions:
                        df_step.at[idx, 'Transmission'] = '\n'.join(all_transmissions)

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя повторить поиск


    #=============== ШАГ 22: ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ BEFORE-ДЕТАЛЕЙ ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ BEFORE-ДЕТАЛЕЙ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # Инициализируем все 18 Before-колонок значением np.nan
        for col in BEFORE_PAC_COLS:
            df_step[col] = np.nan

        # Собираем уникальные Before-детали с их индексами
        before_parts, before_names = collect_unique_parts(
            df_step,
            part_col='Part No. Before',
            name_col='Part Name Before'
        )

        # Проверка на отсутствие Before-деталей
        if not before_parts:
            print(f"\n  [Информация] В BP нет Before-деталей. Шаг {counter} пропущен.")
            print("  Все Before-колонки заполнены значением np.nan (нет данных).")

            show_dataframe_preview(
                df_step,
                step_name=step_name,
                focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
            )

            saved_state = save_state(df_step)

            continue_flag, df_step, saved_state = confirm_step(
                step_name=step_name,
                df_current=df_step,
                saved_state=saved_state
            )

            if continue_flag:
                df_matrix = df_step
                step_elapsed = time.time() - step_start
                print(f"    [Пропуск] Шаг {counter} выполнен за {timedelta(seconds=int(step_elapsed))}")
                break
            continue

        total_unique = len(before_parts)
        packing_dir = INPUT_PACKING_LIST_DIR

        # Выводим список уникальных Before-деталей
        print(f"\n  Всего уникальных Before-деталей: {total_unique}")
        print(f"  Директория упаковочных листов: '{packing_dir}'")
        print("-" * 60)
        print("  СПИСОК BEFORE-ДЕТАЛЕЙ ДЛЯ ОБРАБОТКИ:")
        for i, (part_no, idx_list) in enumerate(sorted(before_parts.items()), 1):
            name = before_names.get(part_no, '')
            name_display = name[:60] + '...' if len(name) > 60 else name
            dup_info = f" [дубликатов: {len(idx_list)}]" if len(idx_list) > 1 else ""
            print(f"    {i:>2}. {part_no} — {name_display}{dup_info}")
        print("-" * 60)
        print("  ИНСТРУКЦИЯ:")
        print("    1. Для каждой уникальной Before-детали найдите её номер партии в SCM.")
        print("    2. Найдите соответствующий упаковочный лист в папке на общем диске.")
        print(f"    3. Скопируйте файл в директорию проекта: '{packing_dir}'")
        print("    4. Введите имя файла упаковочного листа (Enter — пропустить деталь).")
        print("    5. Затем для каждой детали введите тип упаковки и штабелирование.")
        print("    6. Пустой ввод (Enter) оставляет поле неопределённым (np.nan).")
        print("-" * 60)

        wait_for_user("\n  Нажмите Enter, чтобы начать ввод...")

        # Результаты по каждой уникальной детали
        found_results: Dict[str, Dict] = {}      # part_no -> данные из упаковочного листа
        not_found_parts: List[str] = []          # детали без упаковочного листа

        # Интерактив по каждой уникальной Before-детали
        for i, part_no in enumerate(sorted(before_parts.keys()), 1):
            part_name = before_names.get(part_no, '')
            print(f"\n  [{i}/{total_unique}] Before-деталь: {part_no}")
            display_name = part_name[:70] + '...' if len(part_name) > 70 else part_name
            print(f"  Название: {display_name}")

            # --- 1. Загрузка упаковочного листа (если есть) ---
            packaging: Optional[Dict[str, Any]] = None
            while True:
                filename = ask_user_input(
                    "  Имя файла упаковочного листа [Enter — пропустить]: "
                ).strip()

                if not filename:
                    print("    → Пропущено. Данные из файла будут np.nan.")
                    not_found_parts.append(part_no)
                    break

                df_batch = load_packing_list_file(filename, packing_dir)
                if df_batch is None:
                    print(f"    [Ошибка] Не удалось загрузить '{filename}'.")
                    print(f"    Проверьте, что файл лежит в '{packing_dir}', и повторите ввод.")
                    continue

                packaging = extract_packaging_data(df_batch, part_no)
                if packaging is None:
                    print(f"    [Внимание] Деталь '{part_no}' не найдена в файле '{filename}'.")
                    print("    Попробуйте другой файл или нажмите Enter, чтобы пропустить.")
                    continue

                found_results[part_no] = packaging
                print(f"    ✓ Данные из файла '{filename}' успешно извлечены.")
                break

            # --- 2. Пользовательский ввод (тип упаковки, штабелирование) ---
            user_inputs = get_packaging_user_inputs(
                part_no=part_no,
                part_name=part_name,
                branch='Before'
            )

            # --- 3. Заполнение всех 18 колонок для этой детали (включая дубликаты) ---
            apply_packaging_to_rows(
                df=df_step,
                parts_map=before_parts,
                part_no=part_no,
                packaging=packaging,
                user_inputs=user_inputs,
                branch='Before'
            )

        # --- Итоговая сводка ---
        print("\n" + "=" * 60)
        print("  ИТОГОВЫЙ СВОД УПАКОВОЧНЫХ ДАННЫХ BEFORE:")
        print("=" * 60)
        for part_no in sorted(before_parts.keys()):
            name = before_names.get(part_no, '')
            name_display = name[:50] + '...' if len(name) > 50 else name

            if part_no in found_results:
                pkg = found_results[part_no]
                qty_display = data_absence_marker(pkg.get('parts_qty_box'))
                box_display = data_absence_marker(pkg.get('box_size'))
                pal_display = data_absence_marker(pkg.get('pallet_size'))
                print(f"  ✓ {part_no} — {name_display}")
                print(f"      Qty/Box={qty_display}, Box={box_display}, Pallet={pal_display}")
            else:
                print(f"  ✗ {part_no} — {name_display}: упаковочный лист не загружен")
        print("=" * 60)

        if not_found_parts:
            print(f"\n  [Внимание] Для {len(not_found_parts)} Before-деталей упаковочный лист не загружен.")
            print("  Данные из файла для них заполнены np.nan (нет данных).")

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя повторить ввод файлов


    #=============== ШАГ 23: ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ AFTER-ДЕТАЛЕЙ ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ AFTER-ДЕТАЛЕЙ"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        df_step = df_matrix.copy()

        # Инициализируем все 18 After-колонок значением np.nan
        for col in AFTER_PAC_COLS:
            df_step[col] = np.nan

        # Собираем уникальные Before- и After-детали с их индексами
        before_parts, _ = collect_unique_parts(
            df_step,
            part_col='Part No. Before',
            name_col='Part Name Before'
        )
        after_parts, after_names = collect_unique_parts(
            df_step,
            part_col='Part No. After',
            name_col='Part Name After'
        )

        # Проверка на отсутствие After-деталей
        if not after_parts:
            print(f"\n  [Информация] В BP нет After-деталей. Шаг {counter} пропущен.")
            print("  Все After-колонки заполнены значением np.nan (нет данных).")

            show_dataframe_preview(
                df_step,
                step_name=step_name,
                focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
            )

            saved_state = save_state(df_step)

            continue_flag, df_step, saved_state = confirm_step(
                step_name=step_name,
                df_current=df_step,
                saved_state=saved_state
            )

            if continue_flag:
                df_matrix = df_step
                step_elapsed = time.time() - step_start
                print(f"    [Пропуск] Шаг {counter} выполнен за {timedelta(seconds=int(step_elapsed))}")
                break
            continue

        total_unique = len(after_parts)

        # === ЭТАП 1: АВТОКОПИРОВАНИЕ BEFORE → AFTER ===
        print("\n" + "=" * 60)
        print("  АВТОКОПИРОВАНИЕ BEFORE → AFTER")
        print("=" * 60)

        copied_parts = copy_packaging_before_to_after(df_step, before_parts, after_parts)

        if copied_parts:
            print(f"\n  Автоматически скопировано для {len(copied_parts)} After-деталей:")
            for part_no in sorted(copied_parts.keys()):
                print(f"    ⇉ {part_no}")
        else:
            print("  Нет совпадающих After-деталей с Before.")

        # Убираем скопированные детали из дальнейшей обработки
        after_parts_to_process = {
            part_no: idx_list
            for part_no, idx_list in after_parts.items()
            if part_no not in copied_parts
        }

        # Словари для результатов
        found_after: Dict[str, Dict] = {}
        not_found_after: List[str] = []

        if not after_parts_to_process:
            print("\n  [Информация] Все After-детали скопированы из Before. Ручной ввод не требуется.")
        else:
            # === ЭТАП 2: ИНТЕРАКТИВНАЯ ЗАГРУЗКА ДЛЯ ОСТАВШИХСЯ AFTER-ДЕТАЛЕЙ ===
            packing_dir = INPUT_PACKING_LIST_DIR

            print("\n" + "=" * 60)
            print(f"  ЗАГРУЗКА УПАКОВКИ ДЛЯ {len(after_parts_to_process)} AFTER-ДЕТАЛЕЙ")
            print("=" * 60)
            print("  СПИСОК AFTER-ДЕТАЛЕЙ ДЛЯ ОБРАБОТКИ:")
            for i, part_no in enumerate(sorted(after_parts_to_process.keys()), 1):
                name = after_names.get(part_no, '')
                name_display = name[:60] + '...' if len(name) > 60 else name
                print(f"    {i:>2}. {part_no} — {name_display}")
            print("-" * 60)
            print("  ИНСТРУКЦИЯ:")
            print("    1. Для каждой After-детали найдите её упаковочный лист.")
            print(f"    2. Скопируйте файл в директорию проекта: '{packing_dir}'")
            print("    3. Введите имя файла упаковочного листа (Enter — пропустить деталь).")
            print("    4. Затем введите тип упаковки и штабелирование.")
            print("    5. Пустой ввод (Enter) оставляет поле неопределённым (np.nan).")
            print("-" * 60)

            wait_for_user("\n  Нажмите Enter, чтобы начать ввод...")

            # Интерактив по каждой уникальной After-детали
            for i, part_no in enumerate(sorted(after_parts_to_process.keys()), 1):
                part_name = after_names.get(part_no, '')
                print(f"\n  [{i}/{len(after_parts_to_process)}] After-деталь: {part_no}")
                display_name = part_name[:70] + '...' if len(part_name) > 70 else part_name
                print(f"  Название: {display_name}")

                # --- 1. Загрузка упаковочного листа (если есть) ---
                packaging: Optional[Dict[str, Any]] = None
                while True:
                    filename = ask_user_input(
                        "  Имя файла упаковочного листа [Enter — пропустить]: "
                    ).strip()

                    if not filename:
                        print("    → Пропущено. Данные из файла будут np.nan.")
                        not_found_after.append(part_no)
                        break

                    df_batch = load_packing_list_file(filename, packing_dir)
                    if df_batch is None:
                        print(f"    [Ошибка] Не удалось загрузить '{filename}'.")
                        print(f"    Проверьте, что файл лежит в '{packing_dir}', и повторите ввод.")
                        continue

                    packaging = extract_packaging_data(df_batch, part_no)
                    if packaging is None:
                        print(f"    [Внимание] Деталь '{part_no}' не найдена в файле '{filename}'.")
                        print("    Попробуйте другой файл или нажмите Enter, чтобы пропустить.")
                        continue

                    found_after[part_no] = packaging
                    print(f"    ✓ Данные из файла '{filename}' успешно извлечены.")
                    break

                # --- 2. Пользовательский ввод (тип упаковки, штабелирование) ---
                user_inputs = get_packaging_user_inputs(
                    part_no=part_no,
                    part_name=part_name,
                    branch='After'
                )

                # --- 3. Заполнение всех 18 колонок для этой детали (включая дубликаты) ---
                apply_packaging_to_rows(
                    df=df_step,
                    parts_map=after_parts_to_process,
                    part_no=part_no,
                    packaging=packaging,
                    user_inputs=user_inputs,
                    branch='After'
                )

        # --- Итоговая сводка ---
        print("\n" + "=" * 60)
        print("  ИТОГОВЫЙ СВОД УПАКОВОЧНЫХ ДАННЫХ AFTER:")
        print("=" * 60)
        for part_no in sorted(after_parts.keys()):
            name = after_names.get(part_no, '')
            name_display = name[:50] + '...' if len(name) > 50 else name

            if part_no in copied_parts:
                print(f"  ⇉ {part_no} — {name_display}: скопировано из Before")
            elif part_no in found_after:
                pkg = found_after[part_no]
                qty_display = data_absence_marker(pkg.get('parts_qty_box'))
                box_display = data_absence_marker(pkg.get('box_size'))
                pal_display = data_absence_marker(pkg.get('pallet_size'))
                print(f"  ✓ {part_no} — {name_display}")
                print(f"      Qty/Box={qty_display}, Box={box_display}, Pallet={pal_display}")
            else:
                print(f"  ✗ {part_no} — {name_display}: упаковочный лист не загружен")
        print("=" * 60)

        if not_found_after:
            print(f"\n  [Внимание] Для {len(not_found_after)} After-деталей упаковочный лист не загружен.")
            print("  Данные из файла для них заполнены np.nan (нет данных).")

        # Обязательный визуальный контроль результатов
        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя повторить ввод файлов


    #=============== ШАГ 24: УПОРЯДОЧИВАНИЕ КОЛОНОК ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] УПОРЯДОЧИВАНИЕ КОЛОНОК"
    print_step_header(counter, bp_total_steps, step_name)

    while True:
        existing_cols = [col for col in BP_DATA_COLUMNS_ORDER if col in df_matrix.columns]
        missing_cols_in_order = sorted(set(BP_DATA_COLUMNS_ORDER) - set(existing_cols))
        extra_cols = sorted(set(df_matrix.columns) - set(BP_DATA_COLUMNS_ORDER))

        print(f"  Выбрано колонок для финального вывода: {len(existing_cols)} из {len(BP_DATA_COLUMNS_ORDER)}")

        if missing_cols_in_order:
            print(f"  Отсутствуют в данных ({len(missing_cols_in_order)}):")
            for col in missing_cols_in_order:
                print(f"    - {col}")

        if extra_cols:
            print(f"  [Внимание] Лишние колонки будут исключены ({len(extra_cols)}):")
            for col in extra_cols:
                print(f"    - {col}")

        df_step = df_matrix[existing_cols].copy()

        show_dataframe_preview(
            df_step,
            step_name=step_name,
            focus_columns=STEP_FOCUS_COLUMNS_MAP[counter - 1]
        )

        saved_state = save_state(df_step)

        continue_flag, df_step, saved_state = confirm_step(
            step_name=step_name,
            df_current=df_step,
            saved_state=saved_state
        )

        if continue_flag:
            df_matrix = df_step
            step_elapsed = time.time() - step_start
            print(f"    [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")
            break
        # При 'retry' цикл запустится заново, позволяя повторить упорядочивание


    #=============== ШАГ 25: СОХРАНЕНИЕ РЕЗУЛЬТАТА ===============
    counter += 1
    step_start = time.time()
    step_name = f"ШАГ {counter}: [{bp_number}] СОХРАНЕНИЕ РЕЗУЛЬТАТА"
    print_step_header(counter, bp_total_steps, step_name)

    bp_dataframe: Dict[str, Any] = {
        'bp_number': bp_number,
        'dataframe': df_matrix,
        'processing_time_seconds': time.time() - bp_start_time
    }

    step_elapsed = time.time() - step_start
    print(f"  [Успех] Шаг {counter} успешно выполнен за {timedelta(seconds=int(step_elapsed))}")

    # Выводим итоговое время обработки BP
    total_time = bp_dataframe['processing_time_seconds']
    print(f"\n  [ИТОГО] ВРЕМЯ ОБРАБОТКИ {bp_number}: {timedelta(seconds=int(total_time))}")
    print(f"  ({(total_time/60):.1f} минут)\n")

    return bp_dataframe

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
"""
Модуль bp_summary - формирование итоговой сводной таблицы

Назначение модуля:
    Обработка данных об изменениях деталей (Breakpoint) из нескольких BP файлов,
    формирование сводной таблицы с информацией о замене деталей (Before/After),
    расчёт количества партий на складе, загрузка данных из упаковочных листов.

Основные функции:
    1. Загрузка и обработка конфигурационного файла (configuration.xlsx)
    2. Поиск и связывание пар деталей Before/After
    3. Расчёт количества партий на складе (Quantity batches in SS)
    4. Загрузка данных из упаковочных листов (коробки, паллеты)
    5. Формирование итоговой Excel-таблицы для рассылки

Требования к входным данным:
    - BP файлы (Excel) с колонками: Part No., Part Name (RUS), Quantity,
      Workcenter No., Workcenter Name, Supplier Name (RUS), Localization,
      Change, Update Type, Status и др.
    - Конфигурационный файл configuration.xlsx с листом 'common',
      содержащим колонки 'BOM Product' и 'Quantity vehicle in batch'
    - Упаковочные листы (Excel) с колонками:
        - '零部件号码' или 'Part No.',
        - '纸箱装入数量' или 'Parts Q\'ty-Box'
        - и т.д.

Выходные данные:
    pd.DataFrame с колонками: BP_No, Status, Batch plan, Batch fact,
    Part No. Before/After, Quantity per Vehicle Before/After,
    Quantity per Box Before/After, Box/Pallet размеры и др.

Использование:
    >>> from bp_summary import main as summary_main

Версия: 1.0
Совместимость: Python 3.12.3+, Pandas 3.0.2+, OpenPyXL 3.1.5+
Поддержка: PLD Engineering Center
Дата создания: 2026-05-14
Лицензия: MIT
Статус: Production
"""

import os
import sys
import re
import warnings
from typing import Dict, List, Optional

import numpy as np
import pandas as pd

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def clean_surrogates(text):
    """
    Удаляет суррогатные символы Unicode из строки.

    Суррогатные символы (коды U+D800-U+DFFF) являются невалидными
    в кодировке UTF-8 и вызывают ошибку UnicodeEncodeError при попытке
    сохранения в Excel через xlsxwriter.

    Такие символы могут появляться в данных при:
    - Чтении файлов с повреждённой кодировкой
    - Обработке данных из разных источников с нестандартным Unicode
    - Некорректной конвертации между кодировками

    Аргументы:
        text (str or any): Строка для очистки или любое другое значение.

    Возвращается:
        str or any: Очищенная строка без суррогатных символов или
                    исходное значение, если аргумент не является строкой.

    Примеры:
        >>> clean_surrogates('Прокладка\udcd0 головки')
        'Прокладка головки'
        >>> clean_surrogates(123)
        123
        >>> clean_surrogates(None)
        None
    """
    if isinstance(text, str):
        return re.sub(r'[\ud800-\udfff]', '', text)
    return text


def clean_string_for_excel(text):
    """
    Выполняет полную очистку строки от проблемных символов для сохранения в Excel.

    Обрабатывает следующие проблемы:
    1. Суррогатные символы Unicode (U+D800-U+DFFF) — удаляются полностью
    2. Символы перевода строки (\\n, \\r) — заменяются на обычные пробелы
    3. Символы табуляции (\\t) — заменяются на обычные пробелы
    4. Множественные пробелы — схлопываются в один
    5. Пробелы в начале и конце строки — удаляются

    Функция является композицией нескольких операций очистки и использует
    clean_surrogates() как первый шаг обработки.

    Аргументы:
        text (str or any): Строка для очистки или любое другое значение.

    Возвращается:
        str or any: Полностью очищенная строка без проблемных символов или
                    исходное значение, если аргумент не является строкой.

    Примеры:
        >>> clean_string_for_excel('Прокладка\\nголовки\\udcd0 блока')
        'Прокладка головки блока'
        >>> clean_string_for_excel('Деталь\\t\\tтест')
        'Деталь тест'
        >>> clean_string_for_excel('  много   пробелов  ')
        'много пробелов'
        >>> clean_string_for_excel(None)
        None
    """
    if not isinstance(text, str):
        return text

    # Удаляем суррогатные символы (невалидные в UTF-8)
    text = clean_surrogates(text)

    # Заменяем переводы строк и табуляции на пробелы
    text = text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')

    # Схлопываем множественные пробелы в один
    text = re.sub(r'\s+', ' ', text)

    # Удаляем пробелы в начале и конце
    text = text.strip()

    return text


def clean_dataframe_strings(df):
    """
    Очищает все строковые колонки DataFrame от проблемных символов.
    
    Применяет функцию clean_string_for_excel() ко всем значениям
    в строковых колонках (dtype='object'). Обрабатывает каждую ячейку
    индивидуально, отслеживая количество внесённых исправлений.

    Функция не модифицирует исходный DataFrame, а возвращает его копию
    с очищенными данными. Числовые колонки и колонки с другими типами
    данных остаются без изменений.

    Аргументы:
        df (pd.DataFrame or None): DataFrame для очистки. Может быть None
                                    или пустым DataFrame.

    Возвращается:
        pd.DataFrame or None: Копия DataFrame с очищенными строковыми данными.
                              Возвращает None, если на входе был None.
                              Возвращает пустой DataFrame, если на входе был
                              пустой DataFrame.

    Примечания:
        - Функция выполняет глубокое копирование DataFrame (df.copy())
        - Выводит информационное сообщение при обнаружении проблемных символов
        - Обрабатывает только колонки с dtype='object' (строковые)
        - Каждая ячейка проверяется на принадлежность к типу str

    Примеры:
        >>> df = pd.DataFrame({'A': ['тест\\nстрока', 'нормально'], 'B': [1, 2]})
        >>> clean_df = clean_dataframe_strings(df)
        [Очистка данных] Исправлено 1 ячеек с проблемными символами
        >>> clean_df['A'].iloc[0]
        'тест строка'
    """
    if df is None or df.empty:
        return df

    df_clean = df.copy()
    cleaned_count = 0

    for col in df_clean.columns:
        # Обрабатываем только строковые/объектные колонки
        if df_clean[col].dtype == 'object':
            for idx in df_clean.index:
                val = df_clean.at[idx, col]
                if isinstance(val, str):
                    cleaned_val = clean_string_for_excel(val)
                    if cleaned_val != val:
                        df_clean.at[idx, col] = cleaned_val
                        cleaned_count += 1

    if cleaned_count > 0:
        print(f"  [Очистка данных] Исправлено {cleaned_count} ячеек с проблемными символами")

    return df_clean


def fill_empty_values(df: pd.DataFrame) -> list:
    """
    Заполняет пустые значения во всех колонках согласно заданным типам.
    - Числовые колонки (известный список): NaN/Inf -> 0
    - Колонки с датами -> строки ГГГГ-ММ-ДД, пустые -> '-'
    - Остальные (строковые) колонки: пустые строки, 'nan', 'None', 'NaT' -> '-'
    Возвращает список колонок, в которых были замены.
    """
    # Явный список числовых колонок (на основе ваших спецификаций)
    numeric_columns = [
        'Quantity', 'Quantity in SS', 'Quantity per Vehicle Before',
        'Quantity per Vehicle After', 'Quantity per Box Before',
        'Quantity per Box After', 'Quantity batches in SS'
    ]
    # Колонки с датами (если встречаются)
    datetime_columns = ['Change Date', 'New Part Available Date']

    columns_with_replacements = []

    # 1. Обработка числовых колонок
    for col in numeric_columns:
        if col in df.columns:
            # Преобразуем в числа (ошибки -> NaN)
            df[col] = pd.to_numeric(df[col], errors='coerce')
            # Замена Inf на 0
            if np.isinf(df[col]).any():
                df[col] = df[col].replace([np.inf, -np.inf], 0)
                columns_with_replacements.append(col)
            # Замена NaN на 0
            if df[col].isna().any():
                df[col] = df[col].fillna(0)
                columns_with_replacements.append(col)

    # 2. Обработка datetime колонок
    for col in datetime_columns:
        if col in df.columns:
            temp = pd.to_datetime(df[col], errors='coerce')
            df[col] = temp.apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '-')
            columns_with_replacements.append(col)

    # 3. Все остальные колонки (строковые)
    for col in df.columns:
        if col in numeric_columns or col in datetime_columns:
            continue
        # Приводим к строке
        series_str = df[col].astype(str)
        empty_mask = (
            (series_str.str.strip() == '') |
            (series_str.str.strip() == 'nan') |
            (series_str.str.strip() == 'None') |
            (series_str.str.strip() == 'NaT')
        )
        if empty_mask.any():
            series_str[empty_mask] = '-'
            df[col] = series_str
            columns_with_replacements.append(col)

    return columns_with_replacements


def wait_for_user(prompt="\nНажмите Enter для продолжения..."):
    """
    Ожидает нажатия клавиши Enter от пользователя.

    Аргументы:
        prompt (str): Текст приглашения к вводу. По умолчанию содержит инструкцию.

    Обрабатывается:
        - KeyboardInterrupt (Ctrl+C) - запрашивает подтверждение перед завершением
        - EOFError - завершает программу при обнаружении конца ввода
    """
    while True:  # Цикл для повторной попытки ввода
        try:
            user_input = input(prompt)
            return user_input
        except KeyboardInterrupt:
            print()  # Переход на новую строку
            while True:
                confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                if confirm == 'да':
                    print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                    sys.exit(0)
                elif confirm == 'нет':
                    print("\nПродолжаем работу...")
                    break  # Выходим из внутреннего цикла и продолжаем внешний
                else:
                    print("Пожалуйста, введите 'да' или 'нет'")
            # После break из внутреннего цикла, продолжаем внешний цикл
            continue
        except EOFError:
            print("\n\nОбнаружен конец ввода. Программа завершена.")
            sys.exit(0)


def print_step_header(
        step_num,
        total_steps,
        description
    ):
    """
    Выводит форматированный заголовок шага обработки.

    Аргументы:
        step_num (int): Номер текущего шага.
        total_steps (int): Общее количество шагов.
        description (str): Описание шага.
    """
    print("\n" + "=" * 60)
    print(f"ШАГ {step_num}/{total_steps}: {description}")
    print("=" * 60)


def save_state_before_step(
        df
    ):
    """
    Сохраняет состояние DataFrame перед выполнением шага для возможности отката.

    Аргументы:
        df (pd.DataFrame): DataFrame для сохранения.

    Возвращается:
        pd.DataFrame: Глубокая копия DataFrame или None, если df=None.
    """
    if df is not None:
        print("  [Сохранено состояние перед шагом]")
        return df.copy(deep=True)
    return None


def restore_state(
        saved_df,
        step_name
    ):
    """
    Восстанавливает сохранённое состояние DataFrame.

    Аргументы:
        saved_df (pd.DataFrame): Сохранённый DataFrame.
        step_name (str): Имя шага для вывода сообщения.

    Возвращается:
        pd.DataFrame: Восстановленный DataFrame или None при ошибке.
    """
    if saved_df is not None:
        print(f"  [Восстанавливаем состояние перед шагом: {step_name}]")
        return saved_df.copy(deep=True)
    print("  Ошибка: нет сохранённого состояния для восстановления!")
    return None


def confirm_step(
        step_name,
        df_current,
        saved_state
    ):
    """
    Запрашивает у пользователя подтверждение после выполнения шага.

    Поддерживает:
        - Enter: подтверждение, сохранение нового состояния
        - 'retry': отмена изменений и повтор шага

    Аргументы:
        step_name (str): Имя шага для вывода сообщений.
        df_current (pd.DataFrame): Текущий DataFrame после выполнения шага.
        saved_state (pd.DataFrame): Сохранённое состояние перед шагом.

    Возвращается:
        tuple: (continue_flag, df, new_saved_state)
            - continue_flag (bool): True если нужно продолжить, False если retry
            - df (pd.DataFrame): Текущий или восстановленный DataFrame
            - new_saved_state (pd.DataFrame): Новое сохранённое состояние
    """
    print(f"\n  Шаг '{step_name}' выполнен.")

    while True:
        try:
            user_input = input(
                "\nПроверьте результат. Если всё корректно, нажмите Enter. Если нужно повторить шаг, введите 'retry': "
            ).strip().lower()

            if user_input == 'retry':
                print(f"Повторяем шаг '{step_name}'...\n")
                restored_df = restore_state(saved_state, step_name)
                if restored_df is not None:
                    return False, restored_df, saved_state
                else:
                    return False, df_current, saved_state
            else:
                print("Продолжаем...\n")
                new_saved_state = save_state_before_step(df_current)
                return True, df_current, new_saved_state

        except KeyboardInterrupt:
            print()
            while True:
                confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                if confirm == 'да':
                    print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                    sys.exit(0)
                elif confirm == 'нет':
                    print("\nПродолжаем работу...")
                    break  # Выходим из внутреннего цикла и продолжаем внешний
                else:
                    print("Пожалуйста, введите 'да' или 'нет'")
            # После break из внутреннего цикла, продолжаем внешний цикл (повторный ввод)
            continue

        except EOFError:
            print("\n\nОбнаружен конец ввода. Программа завершена.")
            sys.exit(0)


def safe_float_convert(
        value,
        default=0.0
    ) -> float:
    """
    Безопасно преобразует значение в float, обрабатывая пустые значения.

    Аргументы:
        value: Значение для преобразования.
        default (float): Значение по умолчанию. По умолчанию 0.0.

    Возвращается:
        float: Преобразованное значение или значение по умолчанию.
    """
    if value is None or value == '' or value == '-':
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def safe_str_convert(
        value,
        default=''
    ) -> str:
    """
    Безопасно преобразует значение в строку, обрабатывая None и ошибки.

    Аргументы:
        value: Значение для преобразования.
        default (str): Значение по умолчанию. По умолчанию ''.

    Возвращается:
        str: Преобразованная строка или значение по умолчанию.
    """
    if value is None:
        return default
    try:
        return str(value).strip()
    except (ValueError, TypeError):
        return default


def load_configuration_file(
        config_filename: str = 'configuration.xlsx'
    ) -> Optional[pd.DataFrame]:
    """
    Загружает конфигурационный файл (лист 'common').

    Конфигурационный файл должен содержать:
        - BOM Product: идентификатор продукта
        - Quantity vehicle in batch: количество автомобилей в партии
        - Batch code: код партии (опционально)
        - Configuration: конфигурация для утилизации (опционально)
        - Transmission: тип трансмиссии (опционально)

    Аргументы:
        config_filename (str): Имя конфигурационного файла. По умолчанию 'configuration.xlsx'.

    Возвращается:
        Optional[pd.DataFrame]: DataFrame с данными конфигурации или None при ошибке.

    Примечания:
        - Файл должен содержать лист с именем 'common'
        - При отсутствии файла выводится предупреждение, но обработка продолжается
    """
    if not os.path.exists(config_filename):
        print(f"Предупреждение: Файл конфигурации '{config_filename}' не найден")
        print("Некоторые поля будут пустыми")
        return None

    try:
        # Загружаем лист 'common'
        df_config = pd.read_excel(config_filename, sheet_name='common')
        print(f"Файл конфигурации '{config_filename}' (лист 'common') загружен: {df_config.shape[0]} строк")

        # Очистка данных от проблемных символов
        df_config = clean_dataframe_strings(df_config)

        # Проверяем наличие необходимых колонок
        required_cols = ['BOM Product', 'Quantity vehicle in batch']
        missing_cols = [col for col in required_cols if col not in df_config.columns]
        if missing_cols:
            print(f"  Ошибка: В листе 'common' отсутствуют обязательные колонки: {missing_cols}")
            print(f"  Доступные колонки: {df_config.columns.tolist()}")
            return None

        return df_config
    except ValueError as e:
        # Ошибка если лист 'common' не найден
        if "Worksheet named 'common' not found" in str(e):
            print(f"Ошибка: В файле '{config_filename}' не найден лист 'common'")
            print("Переименуйте лист с конфигурацией в 'common'")
        else:
            print(f"Ошибка при загрузке файла конфигурации: {e}")
        return None
    except FileNotFoundError:
        print(f"Ошибка: Файл конфигурации '{config_filename}' не найден")
        return None
    except PermissionError:
        print(f"Ошибка: Нет прав для чтения файла '{config_filename}'")
        return None
    except Exception as e:
        print(f"Непредвиденная ошибка при загрузке: {e}")
        return None


def show_dataframe_preview(
        df, step_name, max_rows=5,
        focus_columns=None,
        max_colwidth=40
    ):
    """
    Отображает первые строки DataFrame для визуального контроля.
    Пустые значения (NaN, None, NaT) показываются как пустая ячейка.
    Числовые значения отображаются как есть (0 не заменяется).
    """
    if df is None or df.empty:
        print(f"\n[Preview после шага: {step_name}]")
        print("  DataFrame пуст!")
        return

    print(f"\n[Preview после шага: {step_name}]")

    # Определяем колонки для отображения
    if focus_columns is None:
        display_cols = df.columns[:5].tolist()
        print(f"Показаны первые 5 колонок из {len(df.columns)}")
    else:
        display_cols = [col for col in focus_columns if col in df.columns]
        if len(df.columns) > len(display_cols):
            print(f"Показаны {len(display_cols)} колонок из {len(df.columns)}")

    if not display_cols:
        print("  Нет колонок для отображения!")
        return

    # Создаём копию для отображения (не изменяем исходный DataFrame)
    preview_df = df[display_cols].head(max_rows).copy()

    # Подготавливаем каждую колонку
    for col in preview_df.columns:
        # Заменяем все варианты пустых значений на пустую строку
        preview_df[col] = preview_df[col].apply(
            lambda x: '' if pd.isna(x) else str(x)
        )
        # Обрезаем длинные строки
        preview_df[col] = preview_df[col].apply(
            lambda x: (x[:max_colwidth] + '…') if len(x) > max_colwidth else x
        )

    # Выводим с помощью pandas, пустые строки отображаются как пустые ячейки
    with pd.option_context(
        'display.max_columns', len(display_cols),
        'display.width', None,
        'display.max_colwidth', max_colwidth,
        'display.show_dimensions', False,
        'display.unicode.east_asian_width', True
    ):
        print(preview_df.to_string(index=False, na_rep=''))
    print()


def classify_row_before_after(
        row: pd.Series
    ) -> str:
    """
    Классифицирует деталь на 'Before', 'After' или 'Unknown'.

    Алгоритм:
        - Приоритет 1: колонка 'Change'
            - 'Before Change' → Before
            - 'After Change' → After
        - Приоритет 2: если Change пустой → Unknown
            (Delete, Add, Replace, Update будут обработаны в find_pairs)

    Аргументы:
        row (pd.Series): Строка DataFrame для классификации.

    Возвращается:
        str: 'Before', 'After' или 'Unknown'.
    """
    change_val = safe_str_convert(row.get('Change', ''))

    # Приоритет 1: явное указание в колонке Change
    if change_val == 'Before Change':
        return 'Before'
    elif change_val == 'After Change':
        return 'After'

    # Приоритет 2: Change пустой → все строки в Unknown
    # (Delete, Add, Replace, Update будут обработаны в find_pairs)
    return 'Unknown'


def create_pair_dict(
        before_row: Optional[pd.Series],
        after_row: Optional[pd.Series]
    ) -> Dict:
    """
    Создаёт словарь с данными для пары деталей Before/After.

    Функция объединяет информацию из строк Before и After деталей,
    заполняя общие поля из любой непустой строки, а специфичные поля
    из соответствующих строк.

    Аргументы:
        before_row (Optional[pd.Series]): Строка с деталью Before или None.
        after_row (Optional[pd.Series]): Строка с деталью After или None.

    Возвращается:
        Dict: Словарь с данными для итоговой таблицы, содержащий:
            - Общие поля: BP_No, Batch plan, BOM Product, Status и др.
            - Поля Before: Part No. Before, Quantity per Vehicle Before и др.
            - Поля After: Part No. After, Quantity per Vehicle After и др.
            - Поля для заполнения позже: Batch fact, Change Date, Comments и др.
    """
    result = {}

    source_row = before_row if before_row is not None else after_row
    if source_row is not None:
        result['BP_No'] = safe_str_convert(source_row.get('BP_No', ''))
        result['Batch plan'] = safe_str_convert(source_row.get('In Stock', ''))

        # New Part Available Date - дата (сохраняем как строку в формате ГГГГ-ММ-ДД)
        new_part_date = source_row.get('New Part Available Date', None)
        result['New Part Available Date'] = safe_str_convert(new_part_date, '')
        result['BOM Product'] = safe_str_convert(source_row.get('BOM Product', ''))
        result['Production Part Disposal'] = safe_str_convert(source_row.get('Production Part Disposal', ''))
        result['Interchangeable'] = safe_str_convert(source_row.get('Interchangeable', ''))
        result['Change Description'] = safe_str_convert(source_row.get('Change Description (RUS)', ''))
        result['Change Solution'] = safe_str_convert(source_row.get('Solution (RUS)', ''))
        result['Color Code'] = safe_str_convert(source_row.get('Color Code', ''))
        result['Color Name (RUS)'] = safe_str_convert(source_row.get('Color Name (RUS)', ''))
        result['Status'] = safe_str_convert(source_row.get('Status', ''))
        qty_ss = source_row.get('Quantity in SS', 0)
        result['Quantity in SS'] = safe_float_convert(qty_ss, 0.0)
    else:
        result['BP_No'] = ''
        result['Batch plan'] = ''
        result['New Part Available Date'] = ''
        result['BOM Product'] = ''
        result['Production Part Disposal'] = ''
        result['Interchangeable'] = ''
        result['Change Description'] = ''
        result['Change Solution'] = ''
        result['Color Code'] = ''
        result['Color Name (RUS)'] = ''
        result['Status'] = ''
        result['Quantity in SS'] = 0.0

    # Before деталь
    if before_row is not None:
        result['Part No. Before'] = safe_str_convert(before_row.get('Part No.', ''))
        result['Part Name Before'] = safe_str_convert(before_row.get('Part Name (RUS)', ''))
        qty_before = before_row.get('Quantity', 0)
        result['Quantity per Vehicle Before'] = safe_float_convert(qty_before, 0.0)
        result['Workcenter No. Before'] = safe_str_convert(before_row.get('Workcenter No.', ''))
        result['Workcenter Name Before'] = safe_str_convert(before_row.get('Workcenter Name', ''))
        result['Supplier Name Before'] = safe_str_convert(before_row.get('Supplier Name (RUS)', ''))
        result['Localization Before'] = safe_str_convert(before_row.get('Localization', ''))
    else:
        result['Part No. Before'] = ''
        result['Part Name Before'] = ''
        result['Quantity per Vehicle Before'] = 0.0
        result['Workcenter No. Before'] = ''
        result['Workcenter Name Before'] = ''
        result['Supplier Name Before'] = ''
        result['Localization Before'] = ''

    # After деталь
    if after_row is not None:
        result['Part No. After'] = safe_str_convert(after_row.get('Part No.', ''))
        result['Part Name After'] = safe_str_convert(after_row.get('Part Name (RUS)', ''))
        qty_after = after_row.get('Quantity', 0)
        result['Quantity per Vehicle After'] = safe_float_convert(qty_after, 0.0)
        result['Workcenter No. After'] = safe_str_convert(after_row.get('Workcenter No.', ''))
        result['Workcenter Name After'] = safe_str_convert(after_row.get('Workcenter Name', ''))
        result['Supplier Name After'] = safe_str_convert(after_row.get('Supplier Name (RUS)', ''))
        result['Localization After'] = safe_str_convert(after_row.get('Localization', ''))
    else:
        result['Part No. After'] = ''
        result['Part Name After'] = ''
        result['Quantity per Vehicle After'] = 0.0
        result['Workcenter No. After'] = ''
        result['Workcenter Name After'] = ''
        result['Supplier Name After'] = ''
        result['Localization After'] = ''

    # Поля, которые будут заполнены позже (с корректными типами данных)
    result['Batch fact'] = ''
    result['Change Date'] = None  # date (будет заполнен позже как строка ГГГГ-ММ-ДД)
    result['Quantity batches in SS'] = 0.0
    result['Configuration for old parts using out'] = ''
    result['Batches for old parts using out'] = ''
    result['Transmission'] = ''
    result['Quantity per Box Before'] = 0.0
    result['Quantity per Box After'] = 0.0
    result['Box Before (L-W-H) mm'] = ''
    result['Box After (L-W-H) mm'] = ''
    result['Pallet Before (L-W-H) mm'] = ''
    result['Pallet After (L-W-H) mm'] = ''
    result['Comments'] = ''

    return result


def find_pairs(
        df_bp: pd.DataFrame
    ) -> List[Dict]:
    """
    Находит и связывает пары деталей Before/After из исходного BP файла.

    Алгоритм работы:
        1. Классифицирует строки по колонке 'Change' (Before Change/After Change)
           и по колонке 'Update Type' (Delete, Add, Replace, Update)
        2. Связывает явные Before/After по Part No. или Part Name
        3. Обрабатывает комбинации Add + Delete (Add=After, Delete=Before)
        4. Обрабатывает комбинации Add + Replace/Update (Add=Before, Replace/Update=After)
        5. Оставшиеся непарные строки добавляются как одиночные

    Аргументы:
        df_bp (pd.DataFrame): Исходный DataFrame с данными BP файла.
            Должен содержать колонки: 'Change', 'Update Type', 'Part No.',
            'Part Name (RUS)' и другие.

    Возвращается:
        List[Dict]: Список словарей, каждый словарь представляет одну пару
            или одиночную деталь с заполненными полями Before/After.
    """
    df_bp = df_bp.copy()

    # Классифицируем все строки
    df_bp['_direction'] = df_bp.apply(classify_row_before_after, axis=1)

    # Разделяем по направлениям
    before_rows = df_bp[df_bp['_direction'] == 'Before'].copy()
    after_rows = df_bp[df_bp['_direction'] == 'After'].copy()
    unknown_rows = df_bp[df_bp['_direction'] == 'Unknown'].copy()

    pairs = []

    # === Обработка явных Before/After ===
    used_before = set()
    used_after = set()

    # Приоритет 1: связывание по Part No.
    for before_idx, before_row in before_rows.iterrows():
        before_part_no = safe_str_convert(before_row.get('Part No.', ''))
        if before_part_no == '' or before_part_no == '-':
            continue

        for after_idx, after_row in after_rows.iterrows():
            if after_idx in used_after:
                continue
            after_part_no = safe_str_convert(after_row.get('Part No.', ''))
            if after_part_no == before_part_no:
                pairs.append(create_pair_dict(before_row, after_row))
                used_before.add(before_idx)
                used_after.add(after_idx)
                break

    # Приоритет 2: связывание по Part Name (RUS)
    for before_idx, before_row in before_rows.iterrows():
        if before_idx in used_before:
            continue
        before_part_name = safe_str_convert(before_row.get('Part Name (RUS)', ''))
        if before_part_name == '' or before_part_name == '-':
            continue

        for after_idx, after_row in after_rows.iterrows():
            if after_idx in used_after:
                continue
            after_part_name = safe_str_convert(after_row.get('Part Name (RUS)', ''))
            if after_part_name == before_part_name:
                pairs.append(create_pair_dict(before_row, after_row))
                used_before.add(before_idx)
                used_after.add(after_idx)
                break

    # Before без пары
    for before_idx, before_row in before_rows.iterrows():
        if before_idx not in used_before:
            pairs.append(create_pair_dict(before_row, None))

    # After без пары
    for after_idx, after_row in after_rows.iterrows():
        if after_idx not in used_after:
            pairs.append(create_pair_dict(None, after_row))

    # === Обработка Unknown строк (Add, Replace, Update) ===
    if unknown_rows.empty:
        return pairs

    # Классифицируем по Update Type
    delete_rows = []
    add_rows = []
    replace_update_rows = []

    for _, row in unknown_rows.iterrows():
        update_type = safe_str_convert(row.get('Update Type', ''))
        if update_type == 'Delete':
            delete_rows.append(row)
        elif update_type == 'Add':
            add_rows.append(row)
        elif update_type in ['Replace', 'Update']:
            replace_update_rows.append(row)

    # Обработка Add + Delete (Add = After, Delete = Before)
    used_delete = set()
    used_add_delete = set()

    for add_idx, add_row in enumerate(add_rows):
        add_part_no = safe_str_convert(add_row.get('Part No.', ''))
        if add_part_no == '' or add_part_no == '-':
            continue

        for del_idx, del_row in enumerate(delete_rows):
            if del_idx in used_delete:
                continue
            del_part_no = safe_str_convert(del_row.get('Part No.', ''))
            if del_part_no == add_part_no:
                pairs.append(create_pair_dict(del_row, add_row))
                used_delete.add(del_idx)
                used_add_delete.add(add_idx)
                break

    for add_idx, add_row in enumerate(add_rows):
        if add_idx in used_add_delete:
            continue
        add_part_name = safe_str_convert(add_row.get('Part Name (RUS)', ''))
        if add_part_name == '' or add_part_name == '-':
            continue

        for del_idx, del_row in enumerate(delete_rows):
            if del_idx in used_delete:
                continue
            del_part_name = safe_str_convert(del_row.get('Part Name (RUS)', ''))
            if del_part_name == add_part_name:
                pairs.append(create_pair_dict(del_row, add_row))
                used_delete.add(del_idx)
                used_add_delete.add(add_idx)
                break

    # Обработка Add + Replace/Update (Add = Before, Replace/Update = After)
    used_replace_update = set()
    used_add_replace = set()

    for add_idx, add_row in enumerate(add_rows):
        if add_idx in used_add_delete or add_idx in used_add_replace:
            continue
        add_part_no = safe_str_convert(add_row.get('Part No.', ''))
        if add_part_no == '' or add_part_no == '-':
            continue

        for ru_idx, ru_row in enumerate(replace_update_rows):
            if ru_idx in used_replace_update:
                continue
            ru_part_no = safe_str_convert(ru_row.get('Part No.', ''))
            if ru_part_no == add_part_no:
                pairs.append(create_pair_dict(add_row, ru_row))
                used_add_replace.add(add_idx)
                used_replace_update.add(ru_idx)
                break

    for add_idx, add_row in enumerate(add_rows):
        if add_idx in used_add_delete or add_idx in used_add_replace:
            continue
        add_part_name = safe_str_convert(add_row.get('Part Name (RUS)', ''))
        if add_part_name == '' or add_part_name == '-':
            continue

        for ru_idx, ru_row in enumerate(replace_update_rows):
            if ru_idx in used_replace_update:
                continue
            ru_part_name = safe_str_convert(ru_row.get('Part Name (RUS)', ''))
            if ru_part_name == add_part_name:
                pairs.append(create_pair_dict(add_row, ru_row))
                used_add_replace.add(add_idx)
                used_replace_update.add(ru_idx)
                break

    # Оставшиеся Delete (Before без After)
    for del_idx, del_row in enumerate(delete_rows):
        if del_idx not in used_delete:
            pairs.append(create_pair_dict(del_row, None))

    # Оставшиеся Add (After без Before)
    for add_idx, add_row in enumerate(add_rows):
        if add_idx not in used_add_delete and add_idx not in used_add_replace:
            pairs.append(create_pair_dict(None, add_row))

    # Оставшиеся Replace/Update (After без Before)
    for ru_idx, ru_row in enumerate(replace_update_rows):
        if ru_idx not in used_replace_update:
            pairs.append(create_pair_dict(None, ru_row))

    return pairs


def user_input_for_single_bp(
        df_current: pd.DataFrame,
        bp_number: str
    ) -> pd.DataFrame:
    """
    Запрашивает у пользователя ввод данных для всего технического изменения (BP).
    ...
    """
    print(f"\n--- Ввод данных для BP {bp_number} ---")
    print(f"Всего строк для обработки: {len(df_current)}")

    df_result = df_current.copy()

    print("\nВведите общие данные для всего технического изменения:")

    # Batch fact
    current_batch_fact = safe_str_convert(df_result.iloc[0].get('Batch fact', '')) if len(df_result) > 0 else ''
    print(f"Текущее Batch fact: {current_batch_fact if current_batch_fact else '(пусто)'}")

    # Защищённый ввод для Batch fact
    while True:
        try:
            batch_fact_input = input("Введите Batch fact (или Enter, чтобы оставить пустым): ").strip()
            break
        except KeyboardInterrupt:
            print()
            while True:
                confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                if confirm == 'да':
                    print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                    sys.exit(0)
                elif confirm == 'нет':
                    print("\nПродолжаем работу...")
                    break
                else:
                    print("Пожалуйста, введите 'да' или 'нет'")

    # Change Date
    current_change_date = safe_str_convert(df_result.iloc[0].get('Change Date', '')) if len(df_result) > 0 else ''
    print(f"Текущее Change Date: {current_change_date if current_change_date else '(пусто)'}")

    # Защищённый ввод для Change Date
    while True:
        try:
            change_date_input = input("Введите Change Date (ГГГГ-ММ-ДД или Enter, чтобы оставить пустым): ").strip()
            break
        except KeyboardInterrupt:
            print()
            while True:
                confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                if confirm == 'да':
                    print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                    sys.exit(0)
                elif confirm == 'нет':
                    print("\nПродолжаем работу...")
                    break
                else:
                    print("Пожалуйста, введите 'да' или 'нет'")

    if batch_fact_input:
        df_result['Batch fact'] = batch_fact_input
        print(f"  Batch fact '{batch_fact_input}' применён ко всем {len(df_result)} строкам")

    if change_date_input:
        df_result['Change Date'] = change_date_input
        print(f"  Change Date '{change_date_input}' применён ко всем {len(df_result)} строкам")

    if not batch_fact_input and not change_date_input:
        print("  Данные не введены. Будут заполнены позже в Excel.")

    return df_result


def config_lookup_for_single_bp(
        df_current: pd.DataFrame,
        df_config: Optional[pd.DataFrame]
    ) -> pd.DataFrame:
    """
    Выполняет поиск данных в конфигурационном файле для каждого BP.

    Функция:
        1. Находит Quantity vehicle in batch по BOM Product
        2. Рассчитывает Quantity batches in SS = Quantity in SS / Quantity vehicle in batch
        3. Ищет Configuration, Batch code, Transmission по Batch fact

    Аргументы:
        df_current (pd.DataFrame): Текущий DataFrame с данными BP.
        df_config (Optional[pd.DataFrame]): DataFrame с конфигурацией из файла
            configuration.xlsx (лист 'common').

    Возвращается:
        pd.DataFrame: DataFrame с заполненными полями:
            - Quantity batches in SS
            - Configuration for old parts using out
            - Batches for old parts using out
            - Transmission
    """
    if df_config is None or df_config.empty:
        print("  Конфигурационный файл не загружен. Пропускаем поиск.")
        return df_current

    df_result = df_current.copy()

    # Проверяем наличие необходимых колонок в конфигурации
    if 'BOM Product' not in df_config.columns or 'Quantity vehicle in batch' not in df_config.columns:
        print("  Ошибка: В конфигурационном файле отсутствуют обязательные колонки 'BOM Product' или 'Quantity vehicle in batch'")
        print(f"  Доступные колонки: {df_config.columns.tolist()}")
        return df_result

    # Приводим BOM Product к строковому типу для поиска
    df_config['BOM Product'] = df_config['BOM Product'].astype(str).str.strip()
    if 'Batch code' in df_config.columns:
        df_config['Batch code'] = df_config['Batch code'].astype(str).str.strip()
    if 'Configuration' in df_config.columns:
        df_config['Configuration'] = df_config['Configuration'].astype(str).str.strip()
    if 'Transmission' in df_config.columns:
        df_config['Transmission'] = df_config['Transmission'].astype(str).str.strip()

    for idx, row in df_result.iterrows():
        # Получаем Part No. Before - только для них нужно считать количество партий
        part_no_before = safe_str_convert(row.get('Part No. Before', ''))

        # Пропускаем строки, где нет Before детали (только After)
        if part_no_before == '' or part_no_before == '-':
            continue

        # Получаем BOM Product для поиска в конфигурации
        bom_product = safe_str_convert(row.get('BOM Product', ''))
        batch_fact = safe_str_convert(row.get('Batch fact', ''))

        # Получаем Quantity in SS
        quantity_in_ss = row.get('Quantity in SS', 0)
        qty_in_ss = safe_float_convert(quantity_in_ss, 0.0)

        if bom_product == '' or bom_product == '-':
            continue

        # Ищем в конфигурации по BOM Product
        config_matches = df_config[df_config['BOM Product'] == bom_product]
        if config_matches.empty:
            continue

        # Берем первое найденное значение (все Quantity vehicle in batch для одного BOM Product одинаковые)
        quantity_vehicle_in_batch = config_matches.iloc[0].get('Quantity vehicle in batch')
        qty_vehicle = safe_float_convert(quantity_vehicle_in_batch, 1.0)

        # Вычисляем Quantity batches in SS только для Before деталей
        if qty_vehicle > 0:
            qty_batches = round(qty_in_ss / qty_vehicle, 2)
            df_result.at[idx, 'Quantity batches in SS'] = qty_batches
        else:
            qty_batches = 0.0

        # Выводим пользователю Part No. Before
        print(f"  {part_no_before}: Количество партий в SS = {qty_batches}")

        # Поиск дополнительной конфигурации по Batch fact
        if batch_fact and batch_fact != '' and batch_fact != '-':
            if 'Batch code' in df_config.columns and 'Configuration' in df_config.columns:
                batch_prefix = batch_fact[:3] if len(batch_fact) >= 3 else batch_fact
                config_match = config_matches[
                    config_matches['Batch code'].str.startswith(batch_prefix, na=False)
                ]
                if not config_match.empty:
                    config_value = safe_str_convert(config_match.iloc[0].get('Configuration', ''))
                    if config_value and config_value != 'nan':
                        df_result.at[idx, 'Configuration for old parts using out'] = config_value

                    batch_code_value = safe_str_convert(config_match.iloc[0].get('Batch code', ''))
                    if batch_code_value and batch_code_value != 'nan':
                        df_result.at[idx, 'Batches for old parts using out'] = batch_code_value

                    transmission_value = safe_str_convert(config_match.iloc[0].get('Transmission', ''))
                    if transmission_value and transmission_value != 'nan':
                        df_result.at[idx, 'Transmission'] = transmission_value

                    print(f"  Найдена конфигурация для {batch_fact}")

    return df_result


def batch_file_loader_for_single_bp(
        df_current: pd.DataFrame
    ) -> pd.DataFrame:
    """
    Загружает упаковочные листы и извлекает данные упаковки для деталей.

    Функция:
        1. Запрашивает у пользователя имя файла упаковочного листа для Before деталей
        2. Загружает файл и извлекает:
            - Quantity per Box (количество деталей в коробке)
            - Box Size (размер коробки)
            - Pallet Size (размер паллеты)
        3. Повторяет процедуру для After деталей (только если указан Batch fact)

    Аргументы:
        df_current (pd.DataFrame): Текущий DataFrame с данными BP.

    Возвращается:
        pd.DataFrame: DataFrame с заполненными полями упаковки:
            - Quantity per Box Before / After
            - Box Before / After (L-W-H) mm
            - Pallet Before / After (L-W-H) mm
    """
    print("\n" + "=" * 60)
    print("ЗАГРУЗКА УПАКОВОЧНОГО ЛИСТА")
    print("=" * 60)
    print(f"\nТекущая директория (куда нужно скопировать файлы): {os.getcwd()}")

    df_result = df_current.copy()

    # === Обработка Before деталей ===
    print("\n[ОБРАБОТКА BEFORE ДЕТАЛЕЙ]")
    print("ИНСТРУКЦИЯ:")
    print("  1. Найдите номер партии в системе SCM по номеру детали")
    print("  2. Найдите в папке упаковочный лист (Excel файл) для найденной партии")
    print("  3. Папка с упаковочными листами находится здесь: \\hmmr_share\LD\Custom Clearance\Поставки\Серийный KD parts")
    print("  4. СКОПИРУЙТЕ нужный файл упаковочного листа в ТЕКУЩУЮ папку (где находится программа)")
    print("  5. Введите имя файла для загрузки данных (или Enter чтобы пропустить)")
    print("-" * 60)

    for idx, row in df_result.iterrows():
        part_no_before = safe_str_convert(row.get('Part No. Before', ''))
        part_name_before = safe_str_convert(row.get('Part Name Before', ''))

        if part_no_before and part_no_before != '' and part_no_before != '-':
            print(f"\n  Деталь Before: {part_no_before}")
            print(f"  Название: {part_name_before[:50] + '...' if len(part_name_before) > 50 else part_name_before}")

            while True:
                # Защищённый ввод имени файла
                while True:
                    try:
                        filename = input("  Введите имя файла упаковочного листа (или Enter чтобы пропустить): ").strip()
                        break
                    except KeyboardInterrupt:
                        print()
                        while True:
                            try:
                                confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                                if confirm == 'да':
                                    print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                                    sys.exit(0)
                                elif confirm == 'нет':
                                    print("\nПродолжаем работу...")
                                    break
                                else:
                                    print("Пожалуйста, введите 'да' или 'нет'")
                            except KeyboardInterrupt:
                                print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                                sys.exit(0)
                        continue

                if filename == '':
                    print("    → Пропущено. Данные будут заполнены позже в Excel.")
                    break

                # Загружаем файл
                df_batch = load_batch_file_by_name(filename, os.getcwd())

                if df_batch is not None:
                    # Извлекаем данные для детали
                    batch_data = extract_packaging_data(df_batch, part_no_before)
                    if batch_data:
                        df_result.at[idx, 'Quantity per Box Before'] = batch_data['parts_qty_box']
                        df_result.at[idx, 'Box Before (L-W-H) mm'] = batch_data['box_size']
                        df_result.at[idx, 'Pallet Before (L-W-H) mm'] = batch_data['pallet_size']
                        print(f"    → Данные загружены: Qty/Box={batch_data['parts_qty_box']}, Box={batch_data['box_size']}, Pallet={batch_data['pallet_size']}")
                    else:
                        print(f"    → Деталь {part_no_before} не найдена в файле {filename}")
                    break
                else:
                    print("    → Файл не найден. Проверьте имя файла и попробуйте снова.")

                    # Защищённый ввод для retry
                    while True:
                        try:
                            retry = input("    Повторить? (Enter - да, 'нет' - пропустить): ").strip().lower()
                            if retry in ('', 'да'):
                                break
                            elif retry == 'нет':
                                break
                            else:
                                print("    Некорректный ввод. Нажмите Enter чтобы повторить, или введите 'нет' чтобы пропустить.")
                        except KeyboardInterrupt:
                            print()
                            while True:
                                try:
                                    confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                                    if confirm == 'да':
                                        print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                                        sys.exit(0)
                                    elif confirm == 'нет':
                                        print("\nПродолжаем работу...")
                                        break
                                    else:
                                        print("Пожалуйста, введите 'да' или 'нет'")
                                except KeyboardInterrupt:
                                    print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                                    sys.exit(0)
                            continue

                    if retry == 'нет':
                        print("    → Пропущено.")
                        break
                    # Если Enter или другое значение - продолжаем цикл (повторный ввод имени файла)

    # === Обработка After деталей (только если есть Batch fact) ===
    print("\n[ОБРАБОТКА AFTER ДЕТАЛЕЙ]")
    print("ИНСТРУКЦИЯ:")
    print("  1. Найдите в папке упаковочный лист (Excel файл) для партии указанной в колонке 'Batch fact'")
    print("  2. Папка с упаковочными листами находится здесь: \\hmmr_share\LD\Custom Clearance\Поставки\Серийный KD parts")
    print("  3. СКОПИРУЙТЕ нужный файл упаковочного листа в ТЕКУЩУЮ папку (где находится программа)")
    print("  4. Введите имя файла для загрузки данных (или Enter чтобы пропустить)")
    print("-" * 60)

    for idx, row in df_result.iterrows():
        part_no_after = safe_str_convert(row.get('Part No. After', ''))
        part_name_after = safe_str_convert(row.get('Part Name After', ''))
        batch_fact = safe_str_convert(row.get('Batch fact', ''))

        if part_no_after and part_no_after != '' and part_no_after != '-':
            if batch_fact and batch_fact != '' and batch_fact != '-':
                print(f"\n  Деталь After: {part_no_after}")
                print(f"  Название: {part_name_after[:50] + '...' if len(part_name_after) > 50 else part_name_after}")
                print(f"  Batch fact: {batch_fact}")

                while True:
                    # Защищённый ввод имени файла
                    while True:
                        try:
                            filename = input("  Введите имя файла упаковочного листа (или Enter чтобы пропустить): ").strip()
                            break
                        except KeyboardInterrupt:
                            print()
                            while True:
                                try:
                                    confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                                    if confirm == 'да':
                                        print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                                        sys.exit(0)
                                    elif confirm == 'нет':
                                        print("\nПродолжаем работу...")
                                        break
                                    else:
                                        print("Пожалуйста, введите 'да' или 'нет'")
                                except KeyboardInterrupt:
                                    print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                                    sys.exit(0)
                            continue

                    if filename == '':
                        print("    → Пропущено. Данные будут заполнены позже в Excel.")
                        break

                    # Загружаем файл
                    df_batch = load_batch_file_by_name(filename, os.getcwd())

                    if df_batch is not None:
                        # Извлекаем данные для детали
                        batch_data = extract_packaging_data(df_batch, part_no_after)
                        if batch_data:
                            df_result.at[idx, 'Quantity per Box After'] = batch_data['parts_qty_box']
                            df_result.at[idx, 'Box After (L-W-H) mm'] = batch_data['box_size']
                            df_result.at[idx, 'Pallet After (L-W-H) mm'] = batch_data['pallet_size']
                            print(f"    → Данные загружены: Qty/Box={batch_data['parts_qty_box']}, Box={batch_data['box_size']}, Pallet={batch_data['pallet_size']}")
                        else:
                            print(f"    → Деталь {part_no_after} не найдена в файле {filename}")
                        break
                    else:
                        print("    → Файл не найден. Проверьте имя файла и попробуйте снова.")

                        # Защищённый ввод для retry
                        while True:
                            try:
                                retry = input("    Повторить? (Enter - да, 'нет' - пропустить): ").strip().lower()
                                if retry in ('', 'да'):
                                    break
                                elif retry == 'нет':
                                    break
                                else:
                                    print("    Некорректный ввод. Нажмите Enter чтобы повторить, или введите 'нет' чтобы пропустить.")
                            except KeyboardInterrupt:
                                print()
                                while True:
                                    try:
                                        confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                                        if confirm == 'да':
                                            print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                                            sys.exit(0)
                                        elif confirm == 'нет':
                                            print("\nПродолжаем работу...")
                                            break
                                        else:
                                            print("Пожалуйста, введите 'да' или 'нет'")
                                    except KeyboardInterrupt:
                                        print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                                        sys.exit(0)
                                continue

                        if retry == 'нет':
                            print("    → Пропущено.")
                            break
                        # Если Enter или другое значение - продолжаем цикл (повторный ввод имени файла)
            else:
                print(f"\n  Деталь After: {part_no_after} - пропущена (не указан Batch fact)")

    print("\n" + "=" * 60)
    print("Завершена загрузка файлов партий.")
    print("=" * 60)

    return df_result

def extract_packaging_data(
        df_batch: pd.DataFrame,
        part_no: str
    ) -> Optional[Dict]:
    """
    Извлекает данные упаковки для указанной детали из загруженного упаковочного листа.

    Функция:
        1. Определяет строку с заголовками колонок (поиск '零部件号码' или 'Part No.')
        2. Устанавливает правильные имена колонок
        3. Ищет строку с указанным Part No.
        4. Извлекает значения:
            - Количество деталей в коробке ('纸箱装入数量' / 'Parts Q\'ty-Box')
            - Размер коробки ('纸箱尺寸' / 'Box Size')
            - Размер паллеты ('包装单元尺寸' / 'Pallet Size')

    Аргументы:
        df_batch (pd.DataFrame): DataFrame с данными упаковочного листа.
        part_no (str): Номер детали для поиска.

    Возвращается:
        Optional[Dict]: Словарь с ключами или None, если деталь не найдена:
            - 'parts_qty_box' (float): Количество деталей в коробке
            - 'box_size' (str): Размер коробки в формате "Д×Ш×В мм"
            - 'pallet_size' (str): Размер паллеты в формате "Д×Ш×В мм"
    """
    if df_batch is None or df_batch.empty:
        return None

    # Очищаем DataFrame: удаляем строки с пропущенными Part No.
    df_clean = df_batch.copy()

    # Определяем строку с заголовками (где есть '零部件号码' или 'Part No.')
    header_row_idx = None
    for idx, row in df_clean.iterrows():
        # Преобразуем строку в список строковых значений
        row_values = [str(val).strip() for val in row.values]
        if '零部件号码' in row_values or 'Part No.' in row_values:
            header_row_idx = idx
            break

    if header_row_idx is not None:
        # Преобразуем индекс в числовую позицию
        # Получаем список всех индексов и находим позицию
        idx_list = df_clean.index.tolist()
        if header_row_idx in idx_list:
            header_pos = idx_list.index(header_row_idx)
        else:
            header_pos = 0

        # Получаем заголовки как список строк
        header_row = df_clean.iloc[header_pos]
        new_columns = []
        for i, col in enumerate(header_row.values):
            if pd.isna(col):
                col_name = f'Unnamed_{i}'
            else:
                col_name = str(col).strip()
            new_columns.append(col_name)

        # Удаляем строку с заголовками и все строки до неё
        df_clean = df_clean.iloc[header_pos + 1:].reset_index(drop=True)

        # Устанавливаем новые колонки
        # Убеждаемся, что количество колонок совпадает
        if len(new_columns) == len(df_clean.columns):
            df_clean.columns = new_columns
        else:
            # Если количество не совпадает, используем только уникальные имена
            unique_columns = []
            for i, col in enumerate(new_columns):
                if col == '' or col == 'nan':
                    col = f'Unnamed_{i}'
                if col in unique_columns:
                    col = f'{col}_{i}'
                unique_columns.append(col)
            # Обрезаем или дополняем список колонок до нужной длины
            if len(unique_columns) > len(df_clean.columns):
                unique_columns = unique_columns[:len(df_clean.columns)]
            elif len(unique_columns) < len(df_clean.columns):
                for i in range(len(unique_columns), len(df_clean.columns)):
                    unique_columns.append(f'Unnamed_{i}')
            df_clean.columns = unique_columns

    # Определяем названия нужных колонок (с поддержкой кириллицы и латиницы)
    col_part_no = None
    col_parts_qty_box = None
    col_box_size = None
    col_pallet_size = None

    for col in df_clean.columns:
        col_str = str(col).strip()
        if '零部件号码' in col_str or 'Part No.' in col_str:
            col_part_no = col
        elif '纸箱装入数量' in col_str or 'Parts Q\'ty-Box' in col_str or 'Parts Qty-Box' in col_str:
            col_parts_qty_box = col
        elif '纸箱尺寸' in col_str or 'Box Size' in col_str or 'Box size' in col_str.lower():
            col_box_size = col
        elif '包装单元尺寸' in col_str or 'Pallet Size' in col_str or 'Pallet size' in col_str.lower():
            col_pallet_size = col

    if col_part_no is None:
        print("  Предупреждение: Не найдена колонка с номерами деталей")
        return None

    # Ищем строку с нужным Part No.
    part_no_str = str(part_no).strip()

    # Приводим колонку с Part No. к строковому типу
    df_clean[col_part_no] = df_clean[col_part_no].astype(str).str.strip()

    # Ищем точное совпадение
    mask = df_clean[col_part_no] == part_no_str
    matching_rows = df_clean[mask]

    if matching_rows.empty:
        return None

    # Берем первую найденную строку
    row = matching_rows.iloc[0]

    # Извлекаем данные
    result = {}

    # Количество деталей в коробке
    if col_parts_qty_box is not None:
        qty = row.get(col_parts_qty_box)
        try:
            if qty is not None:
                if isinstance(qty, (int, float)):
                    result['parts_qty_box'] = float(qty)
                elif str(qty).strip() not in ['', '-', 'nan', 'None']:
                    result['parts_qty_box'] = float(qty)
                else:
                    result['parts_qty_box'] = 0.0
            else:
                result['parts_qty_box'] = 0.0
        except (ValueError, TypeError):
            result['parts_qty_box'] = 0.0
    else:
        result['parts_qty_box'] = 0.0

    # Размер коробки
    if col_box_size is not None:
        box_size = row.get(col_box_size)
        if pd.isna(box_size) or str(box_size).strip() == '-':
            result['box_size'] = ''
        else:
            result['box_size'] = str(box_size).strip()
    else:
        result['box_size'] = ''

    # Размер паллеты
    if col_pallet_size is not None:
        pallet_size = row.get(col_pallet_size)
        if pd.isna(pallet_size) or str(pallet_size).strip() == '-':
            result['pallet_size'] = ''
        else:
            result['pallet_size'] = str(pallet_size).strip()
    else:
        result['pallet_size'] = ''

    return result


def load_batch_file_by_name(
        filename: str,
        search_path: str
    ) -> Optional[pd.DataFrame]:
    """
    Загружает упаковочный лист по имени файла из указанной директории.

    Функция:
        1. Проверяет наличие расширения (.xlsx/.xls), добавляет .xlsx по умолчанию
        2. Формирует полный путь к файлу
        3. Загружает первый лист Excel файла

    Аргументы:
        filename (str): Имя файла (с расширением или без).
        search_path (str): Путь к директории для поиска файла.

    Возвращается:
        Optional[pd.DataFrame]: DataFrame с данными упаковочного листа
                                или None при ошибке (файл не найден, нет прав, повреждён).
    """
    if not filename or filename == '':
        print("  Ошибка: Имя файла не указано")
        return None

    # Если расширение не указано, добавляем .xlsx
    if not filename.endswith(('.xlsx', '.xls')):
        filename = filename + '.xlsx'

    full_path = os.path.join(search_path, filename)

    if not os.path.exists(full_path):
        print(f"  Файл не найден: {full_path}")
        return None

    try:
        # Загружаем первый лист
        df_batch = pd.read_excel(full_path, sheet_name=0)
        print(f"  Файл загружен: {filename} ({df_batch.shape[0]} строк)")
        df_batch = clean_dataframe_strings(df_batch)  # Очистка данных от проблемных символов
        return df_batch
    except FileNotFoundError:
        print(f"  Ошибка: Файл '{full_path}' не найден")
        return None
    except PermissionError:
        print(f"  Ошибка: Нет прав для чтения файла '{full_path}'")
        print("  Закройте файл, если он открыт в Excel, и попробуйте снова.")
        return None
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
        print(f"  Ошибка: Файл '{filename}' повреждён или имеет неверный формат: {e}")
        return None
    except ValueError as e:
        if "Excel file format cannot be determined" in str(e):
            print(f"  Ошибка: Не удалось определить формат файла '{filename}'")
            print("  Убедитесь, что файл имеет расширение .xlsx или .xls")
        else:
            print(f"  Ошибка при загрузке файла '{filename}': {e}")
        return None
    except Exception as e:
        print(f"  Непредвиденная ошибка при загрузке файла '{filename}': {e}")
        print(f"  Тип ошибки: {type(e).__name__}")
        return None


def process_single_bp(
        bp_number: str,
        df_bp: pd.DataFrame,
        df_config: Optional[pd.DataFrame],
        interactive: bool = True
    ) -> pd.DataFrame:
    """
    Выполняет полную обработку одного BP файла.

    Последовательность шагов:
        1. Предварительный просмотр строк с явным указанием Change
        2. Поиск и связывание пар Before/After
        3. Ввод данных Batch fact и Change Date (интерактивно)
        4. Поиск данных в конфигурационном файле и расчёт партий
        5. Загрузка данных из упаковочных листов

    Аргументы:
        bp_number (str): Номер BP для вывода в сообщениях.
        df_bp (pd.DataFrame): Исходный DataFrame с данными BP файла.
        df_config (Optional[pd.DataFrame]): DataFrame с конфигурацией.
        interactive (bool): Флаг интерактивного режима. По умолчанию True.
                            Если False, пропускает шаг ввода данных.

    Возвращается:
        pd.DataFrame: Обработанный DataFrame для данного BP или пустой DataFrame
                      при отсутствии данных.
    """
    print(f"\n{'=' * 60}")
    print(f"ОБРАБОТКА BP: {bp_number}")
    print(f"{'=' * 60}")

    # === ПРЕДВАРИТЕЛЬНЫЙ ПРОСМОТР СТРОК С Change ===
    print("\n[ПРЕДВАРИТЕЛЬНЫЙ ПРОСМОТР]")
    print("Строки с явным указанием 'Before Change' / 'After Change' в колонке 'Change':")

    if 'Change' in df_bp.columns:
        # Отбираем строки, где Change содержит 'Before Change' или 'After Change'
        change_mask = df_bp['Change'].astype(str).str.contains('Before Change|After Change', na=False)
        change_rows = df_bp[change_mask].copy()

        if not change_rows.empty:
            print(f"Найдено строк с указанием Change: {len(change_rows)}")

            # Колонки для отображения
            preview_columns = [
                'BP_No', 'BOM Product', 'Change', 'Update Type', 'Is in BOM',
                'Part No.', 'Part Name (RUS)', 'Workcenter No.', 'Workcenter Name',
                'Color Code', 'Color Name (RUS)'
            ]

            # Фильтруем только существующие колонки
            existing_preview_cols = [col for col in preview_columns if col in change_rows.columns]

            if existing_preview_cols:
                show_dataframe_preview(
                    change_rows,
                    "Строки с Before Change/After Change",
                    focus_columns=existing_preview_cols,
                    max_rows=len(change_rows)  # Показываем все найденные строки
                )
            else:
                print("  Нет доступных колонок для отображения preview")
        else:
            print("  Не найдено строк с указанием 'Before Change' или 'After Change'")
            print("  Все изменения будут определяться по колонке 'Update Type'")
    else:
        print("  ВНИМАНИЕ: Колонка 'Change' отсутствует в данных!")
        print("  Все изменения будут определяться только по колонке 'Update Type'")

    print("\n" + "-" * 60)
    wait_for_user("\nНажмите Enter для продолжения поиска пар...")

    saved_state = None
    df_current = None

    # Шаг 1: Поиск пар Before/After
    while True:
        print_step_header(1, 5, "Поиск пар Before/After")
        pairs = find_pairs(df_bp)
        print(f"  Найдено пар/строк: {len(pairs)}")

        if not pairs:
            print(f"  Предупреждение: Для {bp_number} не найдено пар!")
            return pd.DataFrame()

        df_current = pd.DataFrame(pairs)
        show_dataframe_preview(
            df_current, "Поиск пар Before/After",
            focus_columns=['BP_No', 'BOM Product', 'Part No. Before', 'Part No. After', 'Part Name Before', 'Part Name After']
        )

        if saved_state is None:
            saved_state = save_state_before_step(df_current)

        continue_flag, df_current, saved_state = confirm_step("Поиск пар Before/After", df_current, saved_state)
        if continue_flag:
            break

    # Шаг 2: Пользовательский ввод
    if interactive:
        while True:
            print_step_header(2, 5, "Ввод данных Batch fact и Change Date")
            df_current = user_input_for_single_bp(df_current, bp_number)
            show_dataframe_preview(
                df_current, "Ввод данных Batch fact и Change Date",
                focus_columns=['BP_No', 'Batch plan', 'New Part Available Date', 'Batch fact', 'Change Date', 'Part No. Before', 'Part No. After']
            )

            continue_flag, df_current, saved_state = confirm_step("Ввод данных Batch fact и Change Date", df_current, saved_state)
            if continue_flag:
                break

    # Шаг 3: Поиск данных в конфигурационном файле
    while True:
        print_step_header(3, 5, "Поиск данных в конфигурационном файле")
        df_current = config_lookup_for_single_bp(df_current, df_config)
        show_dataframe_preview(
            df_current, "Поиск данных в конфигурационном файле",
            focus_columns=[
                'BP_No', 'BOM Product', 'Part No. Before', 'Quantity in SS',\
                'Quantity batches in SS', 'Configuration for old parts using out', 
                'Batches for old parts using out', 'Transmission',
            ]
        )

        continue_flag, df_current, saved_state = confirm_step("Поиск данных в конфигурационном файле", df_current, saved_state)
        if continue_flag:
            break

    # Шаг 4: Загрузка данных из упаковочного листа
    while True:
        print_step_header(4, 5, "Загрузка данных из упаковочного листа")
        df_current = batch_file_loader_for_single_bp(df_current)
        show_dataframe_preview(
            df_current, "Загрузка данных из упаковочного листа",
            focus_columns=[
                'BP_No', 'Part No. Before', 'Part No. After',
                'Quantity per Box Before', 'Quantity per Box After',
                'Box Before (L-W-H) mm', 'Box After (L-W-H) mm',
                'Pallet Before (L-W-H) mm', 'Pallet After (L-W-H) mm'
            ]
        )

        continue_flag, df_current, saved_state = confirm_step("Загрузка данных из упаковочного листа", df_current, saved_state)
        if continue_flag:
            break

    # Шаг 5: Заполнение пустых значений
    while True:
        print_step_header(5, 5, "Заполнение пустых значений")

        # Заполняем пустые значения во ВСЕХ колонках
        columns_with_replacements = fill_empty_values(df_current)

        if columns_with_replacements:
            print(f"\n  Колонок с пустыми значениями: {len(columns_with_replacements)}")
            # Показываем только колонки, где были замены
            show_dataframe_preview(
                df_current,
                "Заполнение пустых значений",
                focus_columns=columns_with_replacements[:8],  # Первые 8 колонок с заменами
                max_rows=5
            )
            if len(columns_with_replacements) > 8:
                print(f"  ... и ещё {len(columns_with_replacements) - 8} колонок с заменами")
        else:
            print("  Все колонки уже заполнены данными.")
            show_dataframe_preview(
                df_current, "Заполнение пустых значений",
                focus_columns=['BP_No', 'Part No. Before', 'Part No. After', 'Batch fact']
            )

        continue_flag, df_current, saved_state = confirm_step("Заполнение пустых значений", df_current, saved_state)
        if continue_flag:
            break

    print(f"\n  BP {bp_number} обработан. Добавлено строк: {len(df_current)}")
    return df_current


def main(
        processed_results: Dict[str, pd.DataFrame],
        interactive: bool = True
    ) -> pd.DataFrame:
    """
    Главная функция формирования итоговой сводной таблицы.

    Функция последовательно обрабатывает все переданные BP файлы,
    объединяет результаты и формирует итоговый DataFrame с упорядоченными колонками.

    Аргументы:
        processed_results (Dict[str, pd.DataFrame]): Словарь, где ключ - номер BP,
            значение - DataFrame с исходными данными BP файла.
        interactive (bool): Флаг интерактивного режима. По умолчанию True.
            Если False, пропускает шаг ввода данных Batch fact и Change Date.

    Возвращается:
        pd.DataFrame: Итоговая сводная таблица со всеми обработанными данными,
            содержащая колонки в заданном порядке.

    Пример использования:
        >>> from bp_summary import main as summary_main
    """
    print("""
        ╔══════════════════════════════════════════════════════════════╗
        ║              ФОРМИРОВАНИЕ ИТОГОВОЙ ТАБЛИЦЫ                   ║
        ║           И СОХРАНЕНИЕ EXCEL ФАЙЛА ДЛЯ РАССЫЛКИ              ║
        ║           В ЧАТ DING TALK 'BREAKPOINT PLD INFO'              ║
        ╚══════════════════════════════════════════════════════════════╝
        """)

    print(f"Найдено BP файлов для обработки: {len(processed_results)}")

    print("\nБудут обработаны следующие шаги для каждого BP файла:")
    print("   1. Поиск пар Before/After деталей (связывание изменений)")
    print("   2. Ввод данных: Batch fact (номер партии) и Change Date (дата изменения)")
    print("   3. Поиск данных в конфигурационном файле (Quantity vehicle in batch)")
    print("   4. Расчёт количества партий на складе Safety Stock (Quantity batches in SS)")
    print("   5. Поиск конфигурации для утилизации старых деталей (Configuration for old parts using out)")
    print("   6. Загрузка упаковочных листов для Before и After деталей")
    print("   7. Извлечение данных упаковки:")
    print("      - Quantity per Box (количество деталей в коробке)")
    print("      - Box Size (размер коробки)")
    print("      - Pallet Size (размер паллеты)")
    print("   8. Формирование итоговой таблицы со всеми колонками")
    print("-" * 60)

    # Загрузка конфигурационного файла (один раз для всех)
    print("\n" + "=" * 60)
    print("ЗАГРУЗКА КОНФИГУРАЦИОННОГО ФАЙЛА")
    print("=" * 60)
    df_config = load_configuration_file('configuration.xlsx')
    wait_for_user()

    # Список для сбора обработанных DataFrame
    all_processed_dfs = []
    bp_list = list(processed_results.items())

    # Последовательная обработка каждого BP
    for i, (bp_number, df_bp) in enumerate(bp_list, 1):
        print(f"\n{'=' * 60}")
        print(f"BP {i}/{len(bp_list)}: {bp_number}")
        print(f"{'=' * 60}")

        df_processed = process_single_bp(bp_number, df_bp, df_config, interactive)

        if not df_processed.empty:
            all_processed_dfs.append(df_processed)

        if i < len(bp_list):
            wait_for_user("\nНажмите Enter для обработки следующего BP...")

    # Объединение всех результатов
    if all_processed_dfs:
        df_summary = pd.concat(all_processed_dfs, ignore_index=True)
    else:
        df_summary = pd.DataFrame()

    # Упорядочивание колонок
    column_order = [
        'BP_No', 'Status', 'Batch plan', 'New Part Available Date', 'Batch fact',
        'Change Date', 'BOM Product', 'Part No. Before', 'Part Name Before',
        'Quantity in SS', 'Quantity batches in SS', 'Configuration for old parts using out',
        'Batches for old parts using out', 'Transmission', 'Part No. After',
        'Part Name After', 'Workcenter No. Before', 'Workcenter Name Before',
        'Workcenter No. After', 'Workcenter Name After', 'Quantity per Vehicle Before',
        'Quantity per Vehicle After', 'Quantity per Box Before', 'Quantity per Box After',
        'Box Before (L-W-H) mm', 'Box After (L-W-H) mm', 'Pallet Before (L-W-H) mm',
        'Pallet After (L-W-H) mm', 'Production Part Disposal', 'Interchangeable',
        'Supplier Name Before', 'Localization Before', 'Supplier Name After',
        'Localization After', 'Change Description', 'Change Solution', 
        'Color Code', 'Color Name (RUS)', 'Comments'
    ]

    for col in column_order:
        if col not in df_summary.columns:
            df_summary[col] = '-'

    df_summary = df_summary[column_order]

    # Итоги
    print("\n" + "=" * 60)
    print("ИТОГОВАЯ ТАБЛИЦА SUMMARY BREAKPOINT")
    print("=" * 60)
    print(f"Обработано BP файлов: {len(all_processed_dfs)}/{len(processed_results)}")
    print(f"Всего строк в итоговой таблице: {len(df_summary)}")
    print(f"Колонок: {len(df_summary.columns)}")

    return df_summary

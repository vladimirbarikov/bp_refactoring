#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
# pylint: disable=import-outside-toplevel
"""
BP Refactoring Tool - Инженерные ETL-инструменты нижнего уровня (Extract, Transform, Load).

Модуль спроектирован как полностью автономный технический «движок» проекта.
Он изолирован от конкретных бизнес-констант, словарей детекции, цветовых схем
оформления или фиксированных списков колонок. Вся управляющая логика и правила
форматирования передаются в функции снаружи (из core.py).

Архитектура модуля разделена на 5 функциональных слоёв:

    1. БЕЗОПАСНОЕ ПРЕОБРАЗОВАНИЕ ТИПОВ (Type Conversion)
       Скалярные safe-конвертеры для атомарной обработки mixed-типов данных
       в ячейках Excel (None, NaN, пустые строки, некорректные типы) без падения.

    2. ОЧИСТКА И ПРЕОБРАЗОВАНИЕ ДАННЫХ (Data Cleaning & Normalization)
       Векторизованная очистка текстового мусора и суррогатов Unicode, вызывающих
       сбои сохранения в xlsxwriter, а также универсальная нормализация датафреймов
       к жёстко заданным типам (int64, float64, datetime).

    3. ПОИСК EXCEL ФАЙЛОВ И ИЗВЛЕЧЕНИЕ ДАННЫХ (Data Extraction)
       Низкоуровневые инструменты сканирования директорий ввода, поиска свежих
       исторических файлов по шаблону 'YYYY-MM-DD_{prefix}' и безопасного чтения
       листов Excel с поддержкой динамической фильтрации колонок.

    4. РАБОТА С DataFrame (DataFrame Operations)
       Хелперы для извлечения уникальных значений, парсинга содержимого скобок
       и фильтрации китайского оригинала из многострочных ячеек описаний.

    5. СОХРАНЕНИЕ ДАННЫХ И ИНФРАСТРУКТУРА ЭКСПОРТА (Data Export)
       Развёртывание структуры каталогов в 'output_files/', изолированное
       сохранение бэкапов и сборка сложноформатированных финальных отчётов
       с помощью xlsxwriter.

Экспортируемые функции (по слоям):
    Слой 1: str_convert, int_convert, float_convert, date_convert
    Слой 2: clean_surrogates, clean_string, clean_strings,
            data_absence_marker, normalize_data
    Слой 3: find_latest_excel_file, read_excel_file, find_bp_files,
            extract_packaging_data
    Слой 4: get_unique_non_empty_values, extract_parentheses_content,
            keep_chinese_lines
    Слой 5: get_daily_report_path, save_backup,
            save_excel_with_formatting, save_processed_dataframe

Правила работы с путями файловой системы:
    - Модуль ориентирован на запуск из точки входа 'main.py' в корне проекта.
    - Все относительные пути внутри функций (например, 'output_files/', 'input_files/')
      вычисляются относительно корневой директории проекта, а не папки 'py_lib/'.

Использование:
    from py_lib.engine.etl import (
        read_excel_file,
        normalize_data,
        save_processed_dataframe,
    )

Версия: 1.0
Совместимость: Python 3.14.4+, Pandas 3.0.3+, OpenPyXL 3.1.5+
Поддержка: PLD Engineering Center
Дата создания: 2026-09-09
Дата изменения: 2026-09-24
Лицензия: MIT
Статус: Production
"""

import os
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Union

import numpy as np
import pandas as pd
from pandas.errors import EmptyDataError, ParserError
from openpyxl.utils.exceptions import InvalidFileException

# ====================================================================
# 1. БЕЗОПАСНОЕ ПРЕОБРАЗОВАНИЕ ТИПОВ (Type Conversion)
# ====================================================================

def str_convert(
    value: Any,
    default: str = ''
) -> str:
    """
    Преобразует значение в str, обрабатывая пустые значения и ошибки.

    Аргументы:
        value: Значение для преобразования.
        default (str): Значение по умолчанию при ошибке. По умолчанию ''.

    Возвращается:
        str: Преобразованная строка или значение по умолчанию.
    """
    if value is None or value == '' or pd.isna(value):
        return default
    try:
        return str(value).strip()
    except (ValueError, TypeError):
        return default


def int_convert(
    value: Any,
    default: int = 0
) -> int:
    """
    Преобразует значение в int, обрабатывая пустые значения и ошибки.

    Аргументы:
        value: Значение для преобразования.
        default (int): Значение по умолчанию. По умолчанию 0.

    Возвращается:
        int: Преобразованное значение или значение по умолчанию.
    """
    if value is None or value == '' or value == '-' or pd.isna(value):
        return default
    try:
        # Сначала преобразуем во float на случай строк вида '10.0'
        return int(float(value))
    except (ValueError, TypeError):
        return default


def float_convert(
    value: Any,
    default: float = 0.0
) -> float:
    """
    Преобразует значение в float, обрабатывая пустые значения.

    Аргументы:
        value: Значение для преобразования.
        default (float): Значение по умолчанию. По умолчанию 0.0.

    Возвращается:
        float: Преобразованное значение или значение по умолчанию.
    """
    if value is None or value == '' or value == '-' or pd.isna(value):
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def date_convert(
    value: Any,
    default: Optional[datetime] = None,
    pattern: str = '%Y-%m-%d'
) -> Optional[datetime]:
    """
    Безопасно преобразует значение в datetime, обрабатывая различные форматы.

    Поддерживает:
    - datetime объекты (возвращаются как есть)
    - строки в заданном формате
    - числовые timestamp'ы

    Аргументы:
        value: Значение для преобразования.
        default: Значение по умолчанию при ошибке. По умолчанию None.
        pattern (str): Формат строки даты. По умолчанию '%Y-%m-%d'.

    Возвращается:
        Optional[datetime]: Объект datetime или значение по умолчанию.
    """
    try:
        if isinstance(value, datetime):
            return value
        if isinstance(value, str):
            return datetime.strptime(value, pattern)
        if isinstance(value, (int, float)):
            return datetime.fromtimestamp(value)
        return default
    except (ValueError, TypeError):
        return default


# ====================================================================
# 2. ОЧИСТКА И ПРЕОБРАЗОВАНИЕ ДАННЫХ (Data Cleaning)
# ====================================================================

def clean_surrogates(
    text: Any
) -> Any:
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
        >>> clean_surrogates('Прокладка\\udcd0 головки')
        'Прокладка головки'
        >>> clean_surrogates(123)
        123
        >>> clean_surrogates(None)
        None
    """
    if isinstance(text, str):
        return re.sub(r'[\ud800-\udfff]', '', text)
    return text


def clean_string(
    text: Any
) -> Any:
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
        >>> clean_string('Прокладка\\nголовки\\udcd0 блока')
        'Прокладка головки блока'
        >>> clean_string('Деталь\\t\\tтест')
        'Деталь тест'
        >>> clean_string('  много   пробелов  ')
        'много пробелов'
        >>> clean_string(None)
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


def clean_strings(
    df: pd.DataFrame
) -> pd.DataFrame:
    """
    Очищает все строковые элементы в колонках типа 'object' от проблемных символов.
    Абсолютно устойчива к смешанным типам данных (числа, даты, NaN) внутри object-колонок.

    Аргументы:
        df (pd.DataFrame): Исходный DataFrame для очистки.

    Возвращается:
        pd.DataFrame: Копия DataFrame с очищенными текстовыми элементами.
    """
    if df is None or df.empty:
        return df

    df_clean = df.copy()
    cleaned_count = 0

    # Выбираем только колонки типа 'object', где потенциально могут быть строки
    object_cols = df_clean.select_dtypes(include=['object']).columns

    for col in object_cols:
        # Применяем clean_string только если в конкретной ячейке лежит строка (str)
        # Если там лежит float, int или None — оставляем как есть
        cleaned_series = df_clean[col].apply(
            lambda val: clean_string(val) if isinstance(val, str) else val
        )

        # Считаем количество реально изменившихся ячеек
        # Используем .notna(), чтобы корректно сравнивать массивы с пропусками
        try:
            changed_mask = (df_clean[col] != cleaned_series) & df_clean[col].notna() & cleaned_series.notna()
            cleaned_count += changed_mask.sum()
        except Exception:
            # Резервный безопасный подсчет на случай специфических объектов
            pass

        # Записываем очищенные данные обратно
        df_clean[col] = cleaned_series

    if cleaned_count > 0:
        print(f"  [Очистка данных] Исправлено {cleaned_count} текстовых ячеек со смешанным типом данных.")

    return df_clean


def data_absence_marker(
    value: Any,
    default: str = '-'
) -> str:
    """
    Форматирует значение ячейки для безопасного отображения в консоли
    или текстовом отчёте. Унифицирует маркер отсутствия данных.

    Правила:
        - None, np.nan, pd.NaT            → default ('-' по умолчанию)
        - пустая строка                   → default
        - строка только из пробелов       → default
        - строковые суррогаты 'nan',
          'None', 'NaT' (без учёта
          регистра и пробелов)            → default
        - остальные значения              → str(value).strip()

    Аргументы:
        value (Any):     Значение для форматирования.
        default (str):   Маркер отсутствия данных. По умолчанию '-'.

    Возвращается:
        str: Строковое представление значения или маркер отсутствия данных.
    """
    if value is None:
        return default

    # Отлавливаем настоящие np.nan, pd.NaT, pd.NA
    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass

    # Приводим к строке и проверяем на текстовые суррогаты
    text = str(value).strip()
    if not text:
        return default
    if text.lower() in ('nan', 'none', 'nat'):
        return default

    return text


def normalize_data(
    df: pd.DataFrame,
    int_columns: Optional[List[str]] = None,
    float_columns: Optional[List[str]] = None,
    datetime_columns: Optional[List[str]] = None
) -> pd.DataFrame:
    """
    Универсально нормализует DataFrame, выполняя первоочередную очистку строк 
    от суррогатов/переносов и приводя колонки к строго заданным типам.

    Аргументы:
        df (pd.DataFrame): Исходный DataFrame для нормализации.
        int_columns (list): Список колонок для приведения к типу int64.
        float_columns (list): Список колонок для приведения к типу float.
        datetime_columns (list): Список колонок для приведения к датам (ГГГГ-ММ-ДД).

    Возвращается:
        pd.DataFrame: Полностью очищенная и типизированная копия DataFrame.
                      Строковые колонки нормализуются через data_absence_marker,
                      все пропуски и текстовые суррогаты (nan, None, NaT) приводятся к '-'
    """
    if df is None or df.empty:
        return df

    # ШАГ 1: Первоочередная глубокая очистка всех строковых элементов
    df_norm = clean_strings(df)

    int_cols = int_columns or []
    float_cols = float_columns or []
    date_cols = datetime_columns or []

    # ШАГ 2: Обработка числовых колонок INT
    for col in int_cols:
        if col in df_norm.columns:
            df_norm[col] = df_norm[col].apply(lambda x: int_convert(x, default=0))
            df_norm[col] = df_norm[col].astype('int64')

    # ШАГ 3: Обработка числовых колонок FLOAT
    for col in float_cols:
        if col in df_norm.columns:
            df_norm[col] = df_norm[col].apply(lambda x: float_convert(x, default=0.0))
            df_norm[col] = df_norm[col].astype(float)
            if np.isinf(df_norm[col]).any():
                df_norm[col] = df_norm[col].replace([np.inf, -np.inf], 0.0)

    # ШАГ 4: Обработка колонок DATETIME
    for col in date_cols:
        if col in df_norm.columns:
            temp_date = df_norm[col].apply(lambda x: date_convert(x, default=None))
            df_norm[col] = temp_date.apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '-')

    # ШАГ 5: Финальная стандартизация прочих (строковых) колонок
    managed_cols = set(int_cols) | set(float_cols) | set(date_cols)

    for col in df_norm.columns:
        if col in managed_cols:
            continue

        # Единая функция нормализации: np.nan, None, '', 'nan', 'None', 'NaT'
        # (в любом регистре) → '-'. Пустая строка → '-'. Остальное → str(value).
        df_norm[col] = df_norm[col].apply(data_absence_marker)

    return df_norm


# ====================================================================
# 3. ПОИСК EXCEL ФАЙЛОВ И ИЗВЛЕЧЕНИЕ ДАННЫХ (Data Extraction)
# ====================================================================

def find_latest_excel_file(
    file_prefix: str,
    file_path: Optional[str] = None
) -> Optional[str]:
    """
    Находит последний Excel файл с указанным префиксом в имени.
    Поддерживает форматы: .xlsx, .XLSX, .xls, .XLS

    Ищет файлы, соответствующие шаблону: YYYY-MM-DD_{file_prefix}[произвольный_суффикс].xlsx
    и возвращает имя самого свежего файла (по дате в имени).

    Аргументы:
        file_prefix (str):          Префикс имени файла для поиска.
                                    Например: 'breakpoint_data', 'breakpoint_report' и т.д.
        file_path (Optional[str]):  Директория для поиска файлов.
                                    Если None, используется текущая директория.

    Возвращается:
        Optional[str]:              Имя самого свежего файла или None, если файлы не найдены.

    Примеры:
        >>> find_latest_excel_file('breakpoint_data', './input_files/input_breakpoint_data_files')
        '2026-05-14_breakpoint_data.xlsx'
        
        >>> find_latest_excel_file('breakpoint_report', './input_files/input_breakpoint_report')
        '2026-09-22_breakpoint_report.xlsx'
    """
    search_path = file_path or os.getcwd()

    if not os.path.exists(search_path):
        print(f"  Ошибка: Директория '{search_path}' не существует!")
        return None

    pattern = rf"[0-9]{{4}}-[0-9]{{2}}-[0-9]{{2}}_{re.escape(file_prefix)}[^.]*\.(xlsx|xls)"
    matching_files = []

    try:
        for filename in os.listdir(search_path):
            if filename.startswith('~$'):  # Игнорируем временные файлы блокировки Excel
                continue
            if re.match(pattern, filename, re.IGNORECASE):
                date_str = filename[:10]
                try:
                    file_date = datetime.strptime(date_str, '%Y-%m-%d')
                    matching_files.append((file_date, filename))
                except ValueError:
                    continue
    except PermissionError:
        print(f"  Ошибка: Нет прав для чтения директории {search_path}")
        return None
    except OSError as e:
        print(f"  Ошибка при доступе к директории {search_path}: {e}")
        return None

    if not matching_files:
        return None

    # Сортируем по дате (от самой поздней к самой ранней) и возвращаем самый свежий
    matching_files.sort(key=lambda x: x[0], reverse=True)
    latest_file = matching_files[0][1]

    print(f"  Найден последний файл: {latest_file} (от {matching_files[0][0].strftime('%d.%m.%Y')})")
    if len(matching_files) > 1:
        print(f"  Всего найдено файлов с историей: {len(matching_files)}")

    return latest_file


def read_excel_file(
    filename: str,
    file_path: Optional[str] = None,
    sheet_name: Union[str, int, None] = None,
    header: Optional[int] = 0,
    skip_rows: Optional[int] = 0,
    cols: Optional[List[str]] = None
) -> Optional[pd.DataFrame]:
    """
    Читает данные из конкретного Excel файла и возвращает DataFrame 
    с очищенными строками и отфильтрованными колонками.

    Аргументы:
        filename (str):             Имя конкретного файла для загрузки (например, '2026-05-15_bom.XLSX').
        file_path (Optional[str]):  Директория, где лежит файл. Если None, используется текущая папка.
        sheet_name:                 Имя листа или его индекс. По умолчанию None (первый лист).
        header (Optional[int]):     Строка для использования в качестве заголовков колонок.
        skip_rows (Optional[int]):  Количество строк для пропуска в начале файла.
        cols (Optional[List[str]]): Список колонок, которые необходимо оставить на выходе.

    Возвращается:
        Optional[pd.DataFrame]:     DataFrame с очищенными строками и отфильтрованными 
                                    колонками или None при ошибке.
    """
    search_path = file_path or os.getcwd()
    full_path = os.path.join(search_path, filename)

    if not os.path.exists(full_path):
        print(f"  Ошибка: Файл '{filename}' не найден по пути '{search_path}'")
        return None

    read_args: dict[str, Any] = {
        'io': full_path,
        'sheet_name': sheet_name,
        'header': header,
        'skiprows': skip_rows
    }

    try:
        df = pd.read_excel(**read_args)
        print(f"  Файл: {filename} загружен успешно!")

        if cols is not None:
            existing_cols = [col for col in cols if col in df.columns]
            missing_cols = [col for col in cols if col not in df.columns]

            if missing_cols:
                print(f"  Внимание: В файле '{filename}' не найдены запрашиваемые колонки: {missing_cols}")

            df = df[existing_cols]

        print(f"  Размер датафрейма к обработке: {df.shape[0]} строк × {df.shape[1]} колонок")

        # Очищаем данные от проблемных символов/суррогатов
        df = clean_strings(df)
        return df

    except PermissionError:
        print(f"  Ошибка: Нет прав для чтения файла '{filename}'. Закройте его, если он открыт в Excel.")
        return None
    except (EmptyDataError, ParserError) as e:
        print(f"  Ошибка: Файл '{filename}' повреждён или имеет неверный формат: {e}")
        return None
    except InvalidFileException:
        print(f"  Ошибка: Файл '{filename}' не является корректным Excel файлом")
        return None
    except Exception as e:
        print(f"  НЕПРЕДВИДЕННАЯ ОШИБКА при загрузке файла '{filename}': {e}")
        return None


def find_bp_files(
    file_prefix: str,
    file_path: Optional[str] = None
) -> List[str]:
    """
    Ищет BP-файлы по шаблону {file_prefix}*.xlsx в указанной директории
    и возвращает отсортированный список имён файлов.

    Временные файлы Excel (~$...) игнорируются.
    Поиск регистронезависимый: файлы 'BP123.xlsx', 'bp123.xlsx'
    и 'Bp123.xlsx' будут найдены одинаково.

    Аргументы:
        file_prefix (str):     Префикс имени BP-файла. По умолчанию 'BP'.
                               Передаётся снаружи из core.py (BP_FILE_PREFIX).
        file_path (str, opt):  Директория для поиска. Если None — используется
                               текущая рабочая директория.

    Возвращается:
        list: Отсортированный список найденных имён BP-файлов.
    """
    search_path = file_path or os.getcwd()

    if not os.path.exists(search_path):
        print(f"  Ошибка: Директория для поиска BP-файлов '{search_path}' не существует!")
        return []

    pattern = rf'^{re.escape(file_prefix)}.*\.xlsx$'
    bp_files: List[str] = []

    try:
        for filename in os.listdir(search_path):
            if filename.startswith('~$'):
                continue
            if re.match(pattern, filename, re.IGNORECASE):
                bp_files.append(filename)
    except PermissionError:
        print("  Ошибка: Нет прав для чтения директории")
        return []
    except OSError as e:
        print(f"  Ошибка при доступе к директории: {e}")
        return []

    bp_files.sort()
    return bp_files


def extract_packaging_data(
    pac_list_df: pd.DataFrame,
    part_no: str
) -> Optional[Dict[str, Union[float, str, None]]]:
    """
    Извлекает данные упаковки для указанной детали из загруженного упаковочного листа.

    Функция:
        1. Определяет строку с заголовками колонок (поиск '零部件号码' или 'Part No.')
        2. Устанавливает правильные имена колонок
        3. Ищет строку с указанным Part No.
        4. Извлекает значения:
            - Количество деталей в коробке ('纸箱装入数量' / "Parts Q'ty-Box")
            - Размер коробки ('纸箱尺寸' / 'Box Size')
            - Размер паллеты ('包装单元尺寸' / 'Pallet Size')
            - Количество ящиков на поддоне ('纸箱数量' / "Box Q'ty")
            - Вес брутто паллеты ('包装单元毛重' / 'pallet G/W')
            - Вес нетто одной детали ('单个零部件净重' / 'Unit net weight')

    Аргументы:
        pac_list_df (pd.DataFrame): DataFrame с данными упаковочного листа.
        part_no (str): Номер детали для поиска.

    Возвращается:
        Optional[Dict]: Словарь с ключами или None, если деталь не найдена:
            - 'parts_qty_box' (float): Количество деталей в коробке (или np.nan)
            - 'box_size' (str): Размер коробки в формате "Д×Ш×В"
            - 'pallet_size' (str): Размер паллеты в формате "Д×Ш×В"
            - 'boxes_per_pallet' (float): Количество ящиков на 1 паллете (или np.nan)
            - 'pallet_gross_weight' (float): Вес брутто паллеты, кг (или np.nan)
            - 'part_net_weight' (float): Вес нетто одной детали, кг (или np.nan)
    """
    if pac_list_df is None or pac_list_df.empty:
        return None

    pac_list_df_copy = pac_list_df.copy()

    # === Блок определения строки заголовков и установки имён колонок ===
    # (оставлен без изменений)

    header_row_idx = None
    for idx, row in pac_list_df_copy.iterrows():
        row_values = [str(val).strip() for val in row.values]
        if '零部件号码' in row_values or 'Part No.' in row_values:
            header_row_idx = idx
            break

    if header_row_idx is not None:
        idx_list = pac_list_df_copy.index.tolist()
        if header_row_idx in idx_list:
            header_pos = idx_list.index(header_row_idx)
        else:
            header_pos = 0

        header_row = pac_list_df_copy.iloc[header_pos]
        new_columns = []
        for i, col in enumerate(header_row.values):
            if pd.isna(col):
                col_name = f'Unnamed_{i}'
            else:
                col_name = str(col).strip()
            new_columns.append(col_name)

        pac_list_df_copy = pac_list_df_copy.iloc[header_pos + 1:].reset_index(drop=True)

        if len(new_columns) == len(pac_list_df_copy.columns):
            pac_list_df_copy.columns = new_columns
        else:
            unique_columns = []
            for i, col in enumerate(new_columns):
                if col == '' or col == 'nan':
                    col = f'Unnamed_{i}'
                if col in unique_columns:
                    col = f'{col}_{i}'
                unique_columns.append(col)
            if len(unique_columns) > len(pac_list_df_copy.columns):
                unique_columns = unique_columns[:len(pac_list_df_copy.columns)]
            elif len(unique_columns) < len(pac_list_df_copy.columns):
                for i in range(len(unique_columns), len(pac_list_df_copy.columns)):
                    unique_columns.append(f'Unnamed_{i}')
            pac_list_df_copy.columns = unique_columns

    # === Определение названий нужных колонок ===
    col_part_no = None
    col_parts_qty_box = None
    col_box_size = None
    col_pallet_size = None
    col_boxes_per_pallet = None
    col_pallet_gross_weight = None   # NEW
    col_part_net_weight = None       # NEW

    for col in pac_list_df_copy.columns:
        col_str = str(col).strip()
        if '零部件号码' in col_str or 'Part No.' in col_str:
            col_part_no = col
        elif '纸箱装入数量' in col_str or 'Parts Q\'ty-Box' in col_str or 'Parts Qty-Box' in col_str:
            col_parts_qty_box = col
        elif '纸箱尺寸' in col_str or 'Box Size' in col_str or 'Box size' in col_str.lower():
            col_box_size = col
        elif '包装单元尺寸' in col_str or 'Pallet Size' in col_str or 'Pallet size' in col_str.lower():
            col_pallet_size = col
        elif '纸箱数量' in col_str or 'Box Q\'ty' in col_str or 'Box Qty' in col_str:
            col_boxes_per_pallet = col
        elif '包装单元毛重' in col_str or 'pallet G/W' in col_str.lower():
            col_pallet_gross_weight = col
        elif '单个零部件净重' in col_str or 'Unit net weight' in col_str.lower():
            col_part_net_weight = col

    if col_part_no is None:
        print("  Предупреждение: Не найдена колонка с номерами деталей")
        return None

    # === Поиск строки с нужным Part No. ===
    part_no_str = str(part_no).strip()
    pac_list_df_copy[col_part_no] = pac_list_df_copy[col_part_no].astype(str).str.strip()

    mask = pac_list_df_copy[col_part_no] == part_no_str
    matching_rows = pac_list_df_copy[mask]

    if matching_rows.empty:
        return None

    row = matching_rows.iloc[0]
    result = {}

    # --- Количество деталей в коробке ---
    if col_parts_qty_box is not None:
        qty = row.get(col_parts_qty_box)
        try:
            if qty is not None and str(qty).strip() not in ['', '-', 'nan', 'None']:
                result['parts_qty_box'] = float(qty)
            else:
                result['parts_qty_box'] = np.nan
        except (ValueError, TypeError):
            result['parts_qty_box'] = np.nan
    else:
        result['parts_qty_box'] = np.nan

    # --- Размер коробки ---
    if col_box_size is not None:
        box_size = row.get(col_box_size)
        if pd.isna(box_size) or str(box_size).strip() == '-':
            result['box_size'] = ''
        else:
            result['box_size'] = str(box_size).strip()
    else:
        result['box_size'] = ''

    # --- Размер паллеты ---
    if col_pallet_size is not None:
        pallet_size = row.get(col_pallet_size)
        if pd.isna(pallet_size) or str(pallet_size).strip() == '-':
            result['pallet_size'] = ''
        else:
            result['pallet_size'] = str(pallet_size).strip()
    else:
        result['pallet_size'] = ''

    # --- Количество ящиков на поддоне ---
    if col_boxes_per_pallet is not None:
        boxes_qty = row.get(col_boxes_per_pallet)
        try:
            if boxes_qty is not None and str(boxes_qty).strip() not in ['', '-', 'nan', 'None']:
                result['boxes_per_pallet'] = float(boxes_qty)
            else:
                result['boxes_per_pallet'] = np.nan
        except (ValueError, TypeError):
            result['boxes_per_pallet'] = np.nan
    else:
        result['boxes_per_pallet'] = np.nan

    # --- Вес брутто паллеты ---
    if col_pallet_gross_weight is not None:
        gw = row.get(col_pallet_gross_weight)
        try:
            if gw is not None and str(gw).strip() not in ['', '-', 'nan', 'None']:
                result['pallet_gross_weight'] = float(gw)
            else:
                result['pallet_gross_weight'] = np.nan
        except (ValueError, TypeError):
            result['pallet_gross_weight'] = np.nan
    else:
        result['pallet_gross_weight'] = np.nan

    # --- Вес нетто одной детали ---
    if col_part_net_weight is not None:
        nw = row.get(col_part_net_weight)
        try:
            if nw is not None and str(nw).strip() not in ['', '-', 'nan', 'None']:
                result['part_net_weight'] = float(nw)
            else:
                result['part_net_weight'] = np.nan
        except (ValueError, TypeError):
            result['part_net_weight'] = np.nan
    else:
        result['part_net_weight'] = np.nan

    return result


# ====================================================================
# 4. РАБОТА С DataFrame (DataFrame Operations)
# ====================================================================

def get_unique_non_empty_values(
    series: pd.Series,
    column_name: str
) -> List[str]:
    """
    Получает уникальные значения из серии, исключая пустые значения и '-'.

    Аргументы:
        series (pd.Series): Серия для анализа.
        column_name (str): Имя колонки для вывода сообщений.

    Возвращается:
        list: Список уникальных непустых значений.
    """
    series = series.astype(str)
    unique_values = series.dropna().unique()
    filtered_values = [v for v in unique_values if v != '-' and v != 'nan' and str(v).strip() != '']

    if len(filtered_values) == 0:
        print(f"  Колонка '{column_name}': все значения равны '-' или пустые. Перевод не требуется.")

    return filtered_values


def extract_parentheses_content(
    text: Any
) -> Any:
    """
    Извлекает содержимое скобок из текста.

    Поддерживает как круглые скобки (), так и китайские （）.

    Аргументы:
        text: Текст для обработки (может быть NaN).

    Возвращается:
        str: Содержимое скобок (объединённое через пробел) или исходный текст.
    """
    if pd.isna(text):
        return text

    text_str = str(text)
    matches = re.findall(r'[（(](.*?)[）)]', text_str)

    if matches:
        return ' '.join(matches)

    return text_str


def keep_chinese_lines(
    text: Optional[str]
) -> Optional[str]:
    """
    Оставляет в тексте только строки, содержащие китайские иероглифы.

    Функция применяется к ячейкам Excel с описаниями технических изменений
    и решениями, где в одной ячейке через переводы строк идут:
        - оригинальный китайский текст (источник истины),
        - английский перевод,
        - русский перевод.

    Поскольку переводы могут искажать оригинал, для дальнейшей работы
    инженера оставляется ТОЛЬКО китайский оригинал.

    Правила обработки:
        - Многострочный текст: оставляются строки с китайскими иероглифами,
          склеиваются обратно через '\\n'.
        - Однострочный текст: китайский оригинал гарантированно идёт в начале
          строки, а перевод — после него. Обрезается всё, что идёт после
          последнего китайского символа.
        - Если китайских символов нет — текст возвращается без изменений.
        - Не-строковые значения (NaN, None, числа) — возвращаются как есть.

    Аргументы:
        text (Optional[str]): Текст для фильтрации (может быть не строкой).

    Возвращается:
        Optional[str]: Текст, содержащий только китайский оригинал,
                       или исходное значение, если оно не строка.

    Примеры:
        >>> keep_chinese_lines('增加焊缝\\nWeld seam added\\n焊缝加强')
        '增加焊缝\\n焊缝加强'
        >>> keep_chinese_lines('增加焊缝 Weld seam added')
        '增加焊缝'
        >>> keep_chinese_lines('Weld seam added')
        'Weld seam added'
        >>> keep_chinese_lines(None)
        None
    """
    if not isinstance(text, str):
        return text

    chinese_pattern = re.compile(r'[\u4e00-\u9fff]')

    # Многострочный текст: оставляем только строки с китайскими иероглифами
    if '\n' in text:
        chinese_lines = [
            line.strip()
            for line in text.split('\n')
            if line.strip() and chinese_pattern.search(line)
        ]
        return '\n'.join(chinese_lines) if chinese_lines else text

    # Однострочный текст: китайский оригинал идёт в начале,
    # отрезаем всё после последнего китайского иероглифа
    last_chinese_pos = -1
    for i, char in enumerate(text):
        if chinese_pattern.match(char):
            last_chinese_pos = i

    if last_chinese_pos != -1:
        return text[:last_chinese_pos + 1]

    return text


# ====================================================================
# 5. СОХРАНЕНИЕ ДАННЫХ (Data Export)
# ====================================================================

def get_daily_report_path(
    file_prefix: str,
    output_dir: str,
    date_str: Optional[str] = None
) -> str:
    """
    Возвращает полный путь к дневному файлу отчёта по известному префиксу
    и директории. Используется вызывающим кодом для проверки существования
    файла до вызова save_processed_dataframe().

    Аргументы:
        file_prefix (str): Префикс файла (например, 'breakpoint_data').
        output_dir (str):  Директория, где лежит/будет лежать файл.
        date_str (str, opt): Строка даты YYYY-MM-DD. Если None — берётся сегодняшняя.

    Возвращается:
        str: Полный путь к файлу.
    """
    if date_str is None:
        date_str = datetime.now().strftime('%Y-%m-%d')
    filename = f"{date_str}_{file_prefix}.xlsx"
    return os.path.join(output_dir, filename)


def save_backup(
    df: pd.DataFrame,
    bp_number: str,
    backup_root: Optional[str] = None
) -> str:
    """
    Сохраняет обработанный BP DataFrame в Excel файл внутри папки с сегодняшней датой.

    Аргументы:
        df (pd.DataFrame):      Обработанный DataFrame для бэкапа.
        bp_number (str):        Номер бизнес-процесса (например, 'BP25010410').
        backup_root (str, opt): Полный путь к корневой директории бэкапов.
                                Если None — используется текущая рабочая директория.
                                Пример: 'output_files/output_backup_files'.

    Возвращается:
        str: Полный путь к сохранённому файлу бэкапа или пустая строка при ошибке.
    """
    # 1. ПРОВЕРКА ТИПА ВХОДНЫХ ДАННЫХ (Fail-Fast)
    if not isinstance(df, pd.DataFrame):
        raise TypeError(
            f"Ожидался DataFrame, но получен {type(df)}."
            "\nПереданный объект не имеет необходимых методов для обработки и сохранения."
        )

    # Целевая корневая директория бэкапов
    target_backup_root = backup_root if backup_root is not None else os.getcwd()

    # Формируем имя папки с сегодняшней датой и вложенную папку по номеру BP
    today_str = datetime.now().strftime('%Y-%m-%d')
    date_folder = os.path.join(target_backup_root, f"{today_str}_processed_bp")
    bp_folder = os.path.join(date_folder, bp_number)

    # Проверяем и создаем всю цепочку директорий (makedirs сделает это безопасно)
    if not os.path.exists(bp_folder):
        os.makedirs(bp_folder)
        print(f"  [Бэкап] Создана папка для сохранения резервных копий документов: {bp_folder}")

    # Формируем имя файла и итоговый полный путь
    filename = f"{today_str}_{bp_number}.xlsx"
    filepath = os.path.join(bp_folder, filename)

    # 2. ОЧИСТКА ДАННЫХ
    df_to_save = clean_strings(df)

    if not isinstance(df_to_save, pd.DataFrame):
        raise TypeError(
            f"Функция clean_strings вернула {type(df_to_save)} вместо DataFrame."
        )

    # 3. БЛОК ТЕХНИЧЕСКИХ ОШИБОК (Файловая система, доступы)
    try:
        df_to_save.to_excel(filepath, index=False)
        print(f"  [Бэкап] Сохранён файл: {filepath}")
        return filepath

    except PermissionError:
        print(f"  [ОШИБКА] Нет прав для записи файла: {filepath}")
        print("  Закройте файл, если он открыт в Excel, и попробуйте снова.")
    except OSError as e:
        print(f"  [ОШИБКА] Ошибка файловой системы при создании бэкапа: {e}")
    except Exception as e:
        print(f"  [ОШИБКА] Не удалось сохранить бэкап для {bp_number}: {e}")

    return ""


def save_excel_with_formatting(
    df_new_data: pd.DataFrame,
    output_filename: str,
    sheet_name: str = 'pivot',
    russian_headers: Optional[List[str]] = None,
    theme_colors: Optional[Dict[str, str]] = None,
    column_width: Optional[Dict[int, int]] = None
) -> bool:
    """
    Сохраняет DataFrame в Excel с форматированием и тремя строками заголовков.

    Структура выходного файла:
        - Строка 0: объединённые ячейки "ДЛЯ КЛАДОВЩИКОВ" (стиль merged_cell_format)
        - Строка 1: английские имена колонок (из df_new_data.columns) со стилем header_format
        - Строка 2: русские переводы заголовков (если переданы) со стилем header_format
        - Строка 3 и далее: строки данных с форматированием в зависимости от колонки

    Аргументы:
        df_new_data (pd.DataFrame): DataFrame с данными (без строк заголовков)
        output_filename (str): Имя выходного файла
        sheet_name (str): Имя листа в Excel. По умолчанию 'pivot'
        russian_headers (Optional[List[str]]): Список русских переводов для второй строки.
        theme_colors (Optional[Dict[str, str]]): Словарь цветов.
            Если None — берётся из py_lib.config.core.EXCEL_THEME_COLORS.
        column_width (Optional[Dict[int, int]]): Словарь ширины колонок.
            Если None — берётся из py_lib.config.core.EXCEL_COLUMN_WIDTHS.

    Возвращается:
        bool: True если сохранение успешно, False при ошибке
    """
    from py_lib.config.core import EXCEL_COLUMN_WIDTHS, EXCEL_THEME_COLORS

    colors = theme_colors or EXCEL_THEME_COLORS
    widths = column_width or EXCEL_COLUMN_WIDTHS

    # Защита от любых NaN/Inf
    df_new_data = df_new_data.replace([np.nan, np.inf, -np.inf], 0)

    try:
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet(sheet_name)

            # === Собираем динамические форматы на основе конфига цветов ===
            warehouse_merged_format = workbook.add_format({
                'font_name': 'Arial', 'font_size': 10, 'bold': True,
                'font_color': colors.get('warehouse_text', 'black'), 
                'bg_color': colors.get('warehouse_bg', '#FDE9D9'),
                'valign': 'vcenter', 'align': 'center', 'text_wrap': True, 'border': 1
            })

            header_format = workbook.add_format({
                'font_name': 'Arial', 'font_size': 10, 'bold': True,
                'font_color': colors.get('header_text', 'white'), 
                'bg_color': colors.get('header_bg', '#0F243E'),
                'valign': 'center', 'align': 'center', 'text_wrap': True, 'border': 1
            })

            data_format = workbook.add_format({
                'font_name': 'Arial', 'font_size': 10, 
                'font_color': 'black', 'bg_color': colors.get('data_bg', '#FFFFFF'), 
                'valign': 'top'
            })

            columns_e_p_format = workbook.add_format({
                'font_name': 'Arial', 'font_size': 10, 'font_color': 'black',
                'bg_color': colors.get('warehouse_bg', '#FDE9D9'), 'valign': 'top'
            })

            columns_ac_ad_format = workbook.add_format({
                'font_name': 'Arial', 'font_size': 10, 'font_color': 'black',
                'bg_color': colors.get('warehouse_bg', '#FDE9D9'), 'valign': 'top'
            })

            wrap_format = workbook.add_format({
                'font_name': 'Arial', 'font_size': 10, 'font_color': 'black',
                'bg_color': colors.get('data_bg', '#FFFFFF'), 'text_wrap': True, 'valign': 'top'
            })

            wrap_e_p_format = workbook.add_format({
                'font_name': 'Arial', 'font_size': 10, 'font_color': 'black',
                'bg_color': colors.get('warehouse_bg', '#FDE9D9'), 'text_wrap': True, 'valign': 'top'
            })

            wrap_ac_ad_format = workbook.add_format({
                'font_name': 'Arial', 'font_size': 10, 'font_color': 'black',
                'bg_color': colors.get('warehouse_bg', '#FDE9D9'), 'text_wrap': True, 'valign': 'top'
            })

            columns_wrap_format = workbook.add_format({
                'font_name': 'Arial', 'font_size': 10, 'font_color': 'black',
                'bg_color': colors.get('data_bg', '#FFFFFF'), 'text_wrap': True, 'valign': 'top'
            })

            # === Настройка ширины колонок из конфигурации ===
            # Если для колонки нет явного правила в EXCEL_COLUMN_WIDTHS, задаем стандартные 25
            for col_num in range(len(df_new_data.columns)):
                width = widths.get(col_num, 25)
                worksheet.set_column(col_num, col_num, width)

            columns_e_p_indices = list(range(4, 16))
            columns_ac_ad_indices = [28, 29]
            columns_wrap_indices = [34, 35, 38]

            # === СТРОКА 0: объединённые ячейки ===
            if len(df_new_data.columns) > 15:
                worksheet.merge_range(0, 4, 0, 15, "ДЛЯ КЛАДОВЩИКОВ", warehouse_merged_format)
            if len(df_new_data.columns) > 29:
                worksheet.merge_range(0, 28, 0, 29, "ДЛЯ КЛАДОВЩИКОВ", warehouse_merged_format)

            for col_num in range(len(df_new_data.columns)):
                # Пропускаем индексы из диапазонов 4..15 и 28..29, так как они объединены
                if col_num in columns_e_p_indices or col_num in columns_ac_ad_indices:
                    continue
                worksheet.write(0, col_num, '')

            # === СТРОКА 1: английские заголовки ===
            for col_num, header in enumerate(df_new_data.columns):
                worksheet.write(1, col_num, header, header_format)

            # === СТРОКА 2: русские заголовки ===
            if russian_headers is not None and len(russian_headers) == len(df_new_data.columns):
                for col_num, rus_header in enumerate(russian_headers):
                    worksheet.write(2, col_num, rus_header, header_format)
            else:
                for col_num in range(len(df_new_data.columns)):
                    worksheet.write(2, col_num, '', header_format)

            # === ДАННЫЕ (начиная со строки 3) ===
            for row_num in range(len(df_new_data)):
                for col_num in range(len(df_new_data.columns)):
                    value = df_new_data.iloc[row_num, col_num]
                    is_long_text = isinstance(value, str) and len(value) > 50
                    excel_row = row_num + 3

                    if col_num in columns_wrap_indices:
                        worksheet.write(excel_row, col_num, value, columns_wrap_format)
                    elif col_num in columns_e_p_indices:
                        if is_long_text:
                            worksheet.write(excel_row, col_num, value, wrap_e_p_format)
                        else:
                            worksheet.write(excel_row, col_num, value, columns_e_p_format)
                    elif col_num in columns_ac_ad_indices:
                        if is_long_text:
                            worksheet.write(excel_row, col_num, value, wrap_ac_ad_format)
                        else:
                            worksheet.write(excel_row, col_num, value, columns_ac_ad_format)
                    else:
                        if is_long_text:
                            worksheet.write(excel_row, col_num, value, wrap_format)
                        else:
                            worksheet.write(excel_row, col_num, value, data_format)

        print(f"  Файл '{output_filename}' сохранён с форматированием (три строки заголовков)")
        return True

    except ImportError as e:
        print(f"  Ошибка импорта xlsxwriter: {e}")
        print("  Установите xlsxwriter: pip install xlsxwriter")
        print("  Сохраняем без форматирования...")
        df_new_data.to_excel(output_filename, index=False)
        return False
    except PermissionError:
        print(f"  Ошибка: Нет прав для записи в файл '{output_filename}'")
        print("  Закройте файл, если он открыт в Excel, и попробуйте снова.")
        return False
    except Exception as e:
        print(f"  Ошибка при сохранении с форматированием: {e}")
        print("  Сохраняем без форматирования...")
        df_new_data.to_excel(output_filename, index=False)
        return False


def save_processed_dataframe(
    df_new_data: pd.DataFrame,
    file_prefix: str = 'breakpoint_data',
    column_translation: Optional[Dict[str, str]] = None,
    int_columns: Optional[List[str]] = None,
    datetime_columns: Optional[List[str]] = None,
    force_separate: bool = False,
    history_dir: Optional[str] = None,
    output_dir: Optional[str] = None
) -> Optional[str]:
    """
    Универсально сохраняет обработанный DataFrame с возможностью объединения 
    с существующими историческими данными в структурированную директорию:
        {output_dir}/{YYYY-MM-DD}_{file_prefix}.xlsx

    Аргументы:
        df_new_data (pd.DataFrame):          Новый датафрейм с очищенными данными для экспорта.
        file_prefix (str):                   Префикс имени файла для поиска истории и сохранения.
                                             По умолчанию 'breakpoint_data'.
        column_translation (dict, optional): Словарь {английское_имя_колонки: русский_заголовок}.
                                             Позиционный список русских заголовков собирается
                                             автоматически в порядке колонок df_new_data.
        int_columns (list, optional):        Список колонок для принудительного приведения к типу int64.
        datetime_columns (list, optional):   Список колонок для нормализации к текстовому формату даты YYYY-MM-DD.
        force_separate (bool):               Флаг принудительной записи данных в изолированный новый файл 
                                             без слияния с историей. По умолчанию False.
        history_dir (str, optional):         Полный путь к директории с существующей историей,
                                             с которой выполняется слияние. Если None — слияние
                                             не выполняется, формируется новый файл.
                                             Пример: 'input_files/input_breakpoint_data_files'.
        output_dir (str, optional):          Полный путь к директории, в которой будет сохранен 
                                             файл отчёта. Если None — используется
                                             текущая рабочая директория (os.getcwd()).
                                             Пример: 'output_files/output_breakpoint_data_files'.

    Возвращается:
        Optional[str]: 
            - Полный путь к успешно сохранённому файлу Excel (str).
            - Строковый маркер "STRUCTURE_MISMATCH", если обнаружено несовпадение колонок 
              между новыми и историческими данными (требуется реакция вызывающего уровня).
            - None при возникновении критической ошибки записи или отсутствии данных.
    """
    # Инициализация параметров и путей
    translations_map = column_translation or {}

    # Формируем целевую директорию из переданных снаружи параметров.
    # Если output_dir не задан — используем текущую рабочую директорию.
    base_dir = output_dir if output_dir is not None else os.getcwd()
    today_str = datetime.now().strftime('%Y-%m-%d')
    target_dir = base_dir

    # Собираем позиционный список русских заголовков в порядке колонок df_new_data.
    # Если колонка отсутствует в словаре — используем её английское имя.
    russian_headers_list: List[str] = [
        translations_map.get(col, col)
        for col in df_new_data.columns
    ]

    # Гарантируем существование целевой директории
    if not os.path.exists(target_dir):
        try:
            os.makedirs(target_dir)
            print(f"  [Экспорт] Создана директория для сохранения отчета: {target_dir}")
        except OSError as e:
            print(f"  [ОШИБКА] Не удалось создать директорию {target_dir}: {e}")
            return None

    # Попытка загрузки последней доступной истории для слияния.
    # Если history_dir не передан — слияние не выполняется.
    df_existing: Optional[pd.DataFrame] = None
    if history_dir:
        df_existing = read_excel_file(file_prefix, file_path=history_dir)

    df_combined = df_new_data

    # Сценарий 1: Принудительное сохранение в отдельный файл по требованию бизнес-логики
    if force_separate:
        print("  [Экспорт] Запущено изолированное сохранение по требованию бизнес-логики.")
        filename = f"{today_str}_{file_prefix}_new.xlsx"
        full_output_path = os.path.join(target_dir, filename)

        success = save_excel_with_formatting(
            df_new_data, full_output_path, russian_headers=russian_headers_list
        )
        return full_output_path if success else None

    # Сценарий 2: Попытка объединения с существующей историей
    if isinstance(df_existing, pd.DataFrame) and not df_existing.empty:
        print(f"  [Экспорт] Обнаружен исторический файл. Слияние: {len(df_existing)} строк + {len(df_new_data)} новых строк")

        if list(df_existing.columns) != list(df_new_data.columns):
            print("  [Экспорт] ВНИМАНИЕ: Обнаружено критическое несовпадение структуры колонок.")
            return "STRUCTURE_MISMATCH"

        df_combined = pd.concat([df_existing, df_new_data], ignore_index=True)
        print(f"  [Экспорт] Данные успешно объединены. Итого строк к нормализации: {len(df_combined)}")
    else:
        print(f"  [Экспорт] История не найдена или пуста. Формируется новый отчет с {len(df_combined)} строками")

    # Техническая очистка и нормализация объединенного массива
    print("  [Экспорт] Выполняется нормализация типов объединённого массива данных...")
    df_combined = normalize_data(df_combined, int_columns=int_columns, datetime_columns=datetime_columns)

    # Стандартное сохранение финального файла отчета
    filename = f"{today_str}_{file_prefix}.xlsx"
    full_output_path = os.path.join(target_dir, filename)

    print(f"  [Экспорт] Сохранение файла Excel с форматированием: {full_output_path}")
    success = save_excel_with_formatting(
        df_combined, full_output_path, russian_headers=russian_headers_list
    )

    return full_output_path if success else None

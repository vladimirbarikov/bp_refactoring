#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
# pylint: disable=import-outside-toplevel
"""
BP Refactoring Tool - Центральный конфигурационный модуль проекта.

Модуль является единым хранилищем всех констант, справочников, матриц
и настроек визуального оформления, используемых пайплайном обработки
технических изменений Breakpoint.

Все константы сгруппированы в 13 тематических разделов:

     1. ПУТИ К ДИРЕКТОРИЯМ ПРОЕКТА
        Входные ('input_files/...') и выходные ('output_files/...') каталоги.

     2. ШАГИ ОБРАБОТКИ ТЕХНИЧЕСКИХ ИЗМЕНЕНИЙ
        Текстовые описания 25 шагов пайплайна process_bp_file().

     3. ПРЕФИКСЫ ФАЙЛОВ ВХОДНЫХ/ВЫХОДНЫХ ДАННЫХ
        Префиксы файлов: BP, breakpoint_data, breakpoint_report, configuration.

     4. КОЛОНКИ ФАЙЛА BREAKPOINT REPORT
        BP_REPORT_KEY_COL, BP_REPORT_COLS — структура входящего отчёта.

     5. КОЛОНКИ ФАЙЛА CONFIGURATION.XLSX
        CONF_TO_KEEP_COLS + типизация (int/float/datetime/str).

     6. ОБЩИЕ КОЛОНКИ BP-ФАЙЛОВ И BREAKPOINT_DATA.XLSX
        COMMON_COLS — словарь соответствий G-BOM → матричные колонки.

     7. КОЛОНКИ ФАЙЛА BP<НОМЕР>.XLSX
        BP_GMOM_TO_KEEP_COLS + типизация полей входящего BP.

     8. КОЛОНКИ ФАЙЛА YYYY-MM-DD_BREAKPOINT_DATA.XLSX
        BP_DATA_*_COLS, BEFORE_PAC_COLS, AFTER_PAC_COLS,
        BP_DATA_TRANSLATION_COLS, BP_DATA_COLUMNS_ORDER.

     9. РЕЗУЛЬТИРУЮЩИЕ КОЛОНКИ ШАГОВ PROCESS_BP_FILE()
        STEP_FOCUS_COLUMNS_MAP — колонки предпоказа для каждого шага.

    10. НАСТРОЙКИ СТИЛЕЙ И ВИЗУАЛЬНОГО ОФОРМЛЕНИЯ EXCEL
        EXCEL_THEME_COLORS, EXCEL_COLUMN_WIDTHS — палитра и ширина колонок.

    11. МАТРИЦА ЦЕХОВ / РАБОЧИХ СТАНЦИЙ
        WORKSHOPS, WORKSHOP_PREFIX_TO_CODE — определение цеха по воркцентру.

    12. МАТРИЦА КОД / ИМЯ МОДЕЛЕЙ
        BOM_PRODUCT_MAP — префикс BOM Product Code → название модели.

    13. СПРАВОЧНИКИ ДЛЯ ИНТЕРАКТИВНОГО ВВОДА
        PACKAGING_TYPE — типы упаковки (returnable / non-returnable).

Правила использования:
    - Модуль импортируется всеми слоями проекта: main.py, processing.py,
      etl.py, find_pairs.py, interaction.py.
    - Никаких функций здесь нет — только данные (fail-fast на этапе импорта).
    - Изменения констант влияют на весь проект, правьте аккуратно.

Использование:
    from py_lib.config.core import (
        BP_DATA_COLUMNS_ORDER,
        EXCEL_THEME_COLORS,
        WORKSHOPS,
    )

Версия: 1.0
Совместимость: Python 3.14.4+, Pandas 3.0.3+, OpenPyXL 3.1.5+
Поддержка: PLD Engineering Center
Дата создания: 2026-05-14
Дата изменения: 2026-09-24
Лицензия: MIT
Статус: Production
"""

from typing import Any, Dict, List


# ========================================================================
# 1. ПУТИ К ДИРЕКТОРИЯМ ПРОЕКТА
# ========================================================================

# --- Входные данные ---
INPUT_BREAKPOINT_REPORT_DIR: str = "input_files/input_breakpoint_report_files"
INPUT_BREAKPOINT_DATA_DIR: str = "input_files/input_breakpoint_data_files"
INPUT_BREAKPOINT_FILES_DIR: str = "input_files/input_breakpoint_files"
INPUT_CONFIGURATION_DIR: str = "input_files/input_configuration_files"
INPUT_PACKING_LIST_DIR: str = "input_files/input_packing_list_files"

# --- Выходные данные ---
OUTPUT_BACKUP_ROOT: str = "output_files/output_backup_files"
OUTPUT_BREAKPOINT_DATA_ROOT: str = "output_files/output_breakpoint_data_files"


# ========================================================================
# 2. ШАГИ ОБРАБОТКИ ТЕХНИЧЕСКИХ ИЗМЕНЕНИЙ
# ========================================================================

BP_STEP_DESCRIPTIONS: List[str] = [
    'ЗАГРУЗКА, АУДИТ И НОРМАЛИЗАЦИЯ ДАННЫХ',
    'РАЗДЕЛЕНИЕ ДЕТАЛЕЙ НА СТАРЫЕ И НОВЫЕ',
    'ОПРЕДЕЛЕНИЕ СТАТУСА ТЕХНИЧЕСКОГО ИЗМЕНЕНИЯ',
    'ОПРЕДЕЛЕНИЕ НАЗВАНИЙ МОДЕЛЕЙ ПО КОДАМ',
    'ВВОД ФАКТИЧЕСКОЙ ПАРТИИ',
    'ВВОД ФАКТИЧЕСКОЙ ДАТЫ',
    'ВВОД КОЛИЧЕСТВА BEFORE-ДЕТАЛЕЙ В SAFETY STOCK',
    'ПЕРЕВОД НАЗВАНИЙ ДЕТАЛЕЙ',
    'ПЕРЕВОД ОФИЦИАЛЬНЫХ НАЗВАНИЙ ПОСТАВЩИКОВ',
    'ВВОД СТАТУСА ЛОКАЛИЗАЦИИ ПОСТАВЩИКОВ',
    'ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ BEFORE-ПОСТАВЩИКОВ',
    'ВВОД ДАННЫХ О МЕСТОПОЛОЖЕНИИ AFTER-ПОСТАВЩИКОВ',
    'ПЕРЕВОД ОПИСАНИЯ К ИЗМЕНЕНИЮ',
    'ПЕРЕВОД РЕШЕНИЯ К ИЗМЕНЕНИЮ',
    'ПЕРЕВОД ЦВЕТА',
    'ОБРАБОТКА ЦЕХОВ И РАБОЧИХ СТАНЦИЙ',
    'ПЕРЕВОД РЕБОВАНИЙ ПО УТИЛИЗАЦИИ СТАРЫХ ДЕТАЛЕЙ',
    'ПЕРЕВОД ТРЕБОВАНИЙ ПО ВЗАИМОЗАМЕНЯЕМОСТИ',
    'ЗАГРУЗКА КОНФИГУРАЦИОННОГО ФАЙЛА',
    'РАСЧЁТ ПАРТИЙ В SAFETY STOCK',
    'ПОИСК КОНФИГУРАЦИИ ДЛЯ УТИЛИЗАЦИИ BEFORE-ДЕТАЛЕЙ',
    'ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ BEFORE-ДЕТАЛЕЙ',
    'ПОИСК УПАКОВОЧНЫХ ДАННЫХ ДЛЯ AFTER-ДЕТАЛЕЙ',
    'УПОРЯДОЧИВАНИЕ КОЛОНОК',
    'СОХРАНЕНИЕ РЕЗУЛЬТАТА'
]


# ====================================================================
# 3. ПРЕФИКСЫ ФАЙЛОВ ВХОДНЫХ/ВЫХОДНЫХ ДАННЫХ
# ====================================================================

# Префикс номера BP и, соответственно, файла BP<номер>.xlsx
BP_FILE_PREFIX: str = 'BP'

# Префикс файла накопленной истории обработанных BP
BP_DATA_PREFIX: str = 'breakpoint_data'

# Префикс файла отчёта по техническим изменениям
BP_REPORT_PREFIX: str = 'breakpoint_report'

# Префикс файла конфигурации
CONFIGURATION_PREFIX: str = 'configuration'


# ========================================================================
# 4. КОЛОНКИ ФАЙЛА Breakpoint+Report_согл_опуб_22.09.2026.xlsx
# ========================================================================

# 1. Имя ключевой колонки с номерами BP (используется в processing.py)
BP_REPORT_KEY_COL: str = 'BP No.'

# 2. Список колонок с необходимыми данными
BP_REPORT_COLS = [
    'BP No.'
]

# ========================================================================
# 5. КОЛОНКИ ФАЙЛА YYYY-mm-dd_configuration.xlsx
# ========================================================================

# 1. Колонки с необходимыми данными
CONF_TO_KEEP_COLS: List[str] = [
    'Batch Code', 'BOM Product', 'Model',
    'Transmission', 'Configuration', 'Quantity Vehicle in Batch'
]

# 2. int колонки
CONF_INT_COLS: List[str] = [
    'Quantity Vehicle in Batch'
]

# 3. float колонки
CONF_FLOAT_COLS: List[str] = []

# 4. datetime колонки
CONF_DATETIME_COLS: List[str] = []

# 5. str колонки
CONF_STR_COLS: List[str] = [
    'Batch Code', 'BOM Product', 'Model',
    'Transmission', 'Configuration'
]


# ========================================================================
# 6. ОБЩИЕ КОЛОНКИ ФАЙЛОВ YYYY-mm-dd_breakpoint_data.xlsx и BP<номер>.xlsx
# ========================================================================

COMMON_COLS: Dict[str, Dict[str, str]] = {
    'BOM Product': {
        'BOM Product Code': 'Код модели'
    },
    'Part No.': {
        'Part No. Before': 'Номер "старой" детали',
        'Part No. After': 'Номер "новой" детали', 
    },
    'Part Name(CHN)': {
        'Part Name Before': 'Название "старой" детали',
        'Part Name After': 'Название "новой" детали',
    },
    'Quantity': {
        'Quantity per Vehicle Before': 'Количество "старых" деталей на 1 авто',
        'Quantity per Vehicle After': 'Количество "новых" деталей на 1 авто',
    },
    'Supplier Name': {
        'Supplier Name Before': 'Название поставщика "старых" деталей',
        'Supplier Name After': 'Название поставщика "новых" деталей',
    },
    'Change Description': {
        'Change Description (RUS)': 'Описание технического изменения',
    },
    'Solution': {
        'Change Solution (RUS)': 'Решение технического изменения',
    },
    'Color Code': {
        'Color Code': 'Код цвета',
    },
    'Color Name': {
        'Color Name (RUS)': 'Название цвета',
    },
    'Production Part Disposal': {
        'Production Part Disposal': 'Использование "старых" деталей в производстве',
    },
    'Interchangeable': {
        'Interchangeable': 'Взаимозаменяемость "старых/новых" деталей',
    },
    'In Stock': {
        'Batch Plan': 'Планируемая партия',
    },
    'New Part Available Date': {
        'New Part Available Date': 'Дата выхода "новой" детали',
    },
    'Workcenter No.': {
        'Workcenter Code Before': 'Код производственной линии "старой" детали',
        'Workcenter Code After': 'Код производственной линии "новой" детали',
    },
    'Workcenter Name': {
        'Workcenter Name Before': 'Название производственной линии "старой" детали',
        'Workcenter Name After': 'Название производственной линии "новой" детали',
    },
}


# ========================================================================
# 7. КОЛОНКИ ФАЙЛА BP<номер>.xlsx
# ========================================================================

# 1. Колонки с необходимыми данными
BP_GMOM_TO_KEEP_COLS: List[str] = [
    'Change', 'BOM Product', 'Update Type', 'Part No.', 'Part Name(CHN)', 'Quantity',
    'Supplier Name', 'Change Description', 'Solution', 'Color Code', 'Color Name',
    'Production Part Disposal', 'Interchangeable', 'In Stock', 'New Part Available Date',
    'Workcenter No.', 'Workcenter Name'
]

# 2. int колонки
BP_GBOM_INT_COLS: List[str] = [
    'Quantity'
]

# 3. float колонки
BP_GBOM_FLOAT_COLS: List[str] = []

# 4. datetime колонки
BP_GBOM_DATETIME_COLS: List[str] = [
    'New Part Available Date'
]

# 5. str колонки
BP_GBOM_STR_COLS: List[str] = [
    'Change', 'BOM Product', 'Update Type', 'Part No.', 'Part Name(CHN)', 'Supplier Name',
    'Change Description', 'Solution', 'Color Code', 'Color Name', 'Production Part Disposal',
    'Interchangeable', 'In Stock', 'New Part Available Date', 'Workcenter No.', 'Workcenter Name'
]


# ========================================================================
# 8. КОЛОНКИ ФАЙЛА YYYY-mm-dd_breakpoint_data.xlsx
# ========================================================================

# 1. Имя ключевой колонки с номерами BP (используется в processing.py)
BP_DATA_KEY_COL: str = 'BP No.'

# 2. int колонки
BP_DATA_INT_COLS: List[str] = [
    'Box Length Before mm', 'Box Width Before mm', 'Box Height Before mm',
    'Box Length After mm', 'Box Width After mm', 'Box Height After mm',
    'Pallet Length Before mm', 'Pallet Width Before mm', 'Pallet Height Before mm',
    'Pallet Length After mm', 'Pallet Width After mm', 'Pallet Height After mm',
    'Boxes per Pallet Before units', 'Boxes per Pallet After units'
]

# 3. float колонки
BP_DATA_FLOAT_COLS: List[str] = [
    'Quantity Parts in SS', 'Quantity Batches in SS', 'Quantity per Vehicle Before',
    'Quantity per Vehicle After', 'Quantity per Box Before', 'Quantity per Box After', 
    'Part Net Weight Before kg', 'Part Net Weight After kg', 'Box Volume Before m3',
    'Box Volume After m3',  'Box Area Before m2', 'Box Area After m2', 'Box Net Weight Before kg',
    'Box Net Weight After kg', 'Pallet Volume Before m3', 'Pallet Volume After m3',
    'Pallet Area Before m2', 'Pallet Area After m2', 'Pallet Gross Weight Before kg',
    'Pallet Gross Weight After kg'
]

# 4. datetime колонки
BP_DATA_DATETIME_COLS: List[str] = [
    'Change Date', 'New Part Available Date'
]

# 5. str колонки
BP_DATA_STR_COLS: List[str] = [
    'BP_No', 'Status', 'Batch Plan', 'Batch Fact', 'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Part Name Before', 'Configuration for Old Parts Using Out',
    'Batches for Old Parts Using Out', 'Transmission', 'Part No. After',
    'Part Name After', 'Box Type Before', 'Box Type After', 'Pallet Type Before',
    'Pallet Type After', 'Workcenter Code Before', 'Workcenter Name Before',
    'Workshop Code Before', 'Workshop Name Before', 'Workcenter Code After',
    'Workcenter Name After', 'Workshop Code After', 'Workshop Name After',
    'Production Part Disposal', 'Interchangeable', 'Supplier Name Before',
    'Supplier Location Before', 'Supplier City Before', 'Supplier Street Before',
    'Supplier Building Before', 'Localization Before', 'Supplier Name After',
    'Supplier Location After', 'Supplier City After', 'Supplier Street After',
    'Supplier Building After', 'Localization After', 'Box Stacking Before units',
    'Box Stacking After units', 'Pallet Stacking Before units',
    'Pallet Stacking After units','Change Description (RUS)', 
    'Change Solution (RUS)', 'Color Code', 'Color Name (RUS)', 'Comments'
]

# 6. Колонки по типу Before/After данных упаковки
BEFORE_PAC_COLS: List[str] = [
    'Quantity per Box Before', 'Box Type Before', 'Box Length Before mm',
    'Box Width Before mm', 'Box Height Before mm', 'Box Area Before m2',
    'Box Volume Before m3', 'Box Net Weight Before kg', 'Box Stacking Before units',
    'Boxes per Pallet Before units', 'Pallet Type Before', 'Pallet Length Before mm',
    'Pallet Width Before mm', 'Pallet Height Before mm', 'Pallet Area Before m2',
    'Pallet Volume Before m3', 'Pallet Gross Weight Before kg', 'Pallet Stacking Before units'
]

AFTER_PAC_COLS: List[str] = [
    'Quantity per Box After', 'Box Type After', 'Box Length After mm',
    'Box Width After mm', 'Box Height After mm', 'Box Area After m2',
    'Box Volume After m3', 'Box Net Weight After kg', 'Box Stacking After units',
    'Boxes per Pallet After units', 'Pallet Type After', 'Pallet Length After mm',
    'Pallet Width After mm', 'Pallet Height After mm', 'Pallet Area After m2',
    'Pallet Volume After m3', 'Pallet Gross Weight After kg', 'Pallet Stacking After units'
]

# 7. Колонки на "английском" : "на русском"
BP_DATA_TRANSLATION_COLS: Dict[str, str] = {
    'BP_No': 'Номер переключения',
    'Status': 'Статус переключения',
    'Batch Plan': 'Планируемая партия',
    'New Part Available Date': 'Дата выхода "новой" детали',
    'Batch Fact': 'Фактическая партия',
    'Change Date': 'Фактическая дата',
    'BOM Product Code': 'Код модели',
    'BOM Product Name': 'Название модели',
    'Part No. Before': 'Номер "старой" детали',
    'Part Name Before': 'Название "старой" детали',
    'Quantity Parts in SS': 'Количество "старых" деталей в Safety Stock',
    'Quantity Batches in SS': 'Количество партий со "старыми" деталями в Safety Stock',
    'Configuration for Old Parts Using Out': 'Конфигурация "старых" деталей',
    'Batches for Old Parts Using Out': 'Партии для использования "старых" деталей',
    'Transmission': 'Привод (трансмиссия)',
    'Part No. After': 'Номер "новой" детали',
    'Part Name After': 'Название "новой" детали',
    'Workcenter Code Before': 'Код производственной линии "старой" детали',
    'Workcenter Name Before': 'Название производственной линии "старой" детали',
    'Workshop Code Before': 'Код цеха "старой" детали',
    'Workshop Name Before': 'Название цеха "старой" детали',
    'Workcenter Code After': 'Код производственной линии "новой" детали',
    'Workcenter Name After': 'Название производственной линии "новой" детали',
    'Workshop Code After': 'Код цеха "новой" детали',
    'Workshop Name After': 'Название цеха "новой" детали',
    'Part Net Weight Before kg': 'Вес "старой" детали кг',
    'Part Net Weight After kg': 'Вес "новой" детали кг',
    'Quantity per Vehicle Before': 'Количество "старых" деталей на 1 авто',
    'Quantity per Vehicle After': 'Количество "новых" деталей на 1 авто',
    'Quantity per Box Before': 'Количество "старых" деталей на 1 ящик',
    'Quantity per Box After': 'Количество "новых" деталей на 1 ящик',
    'Box Type Before': 'Тип ящика "старой" детали',
    'Box Length Before mm': 'Длина ящика "старой" детали мм',
    'Box Width Before mm': 'Ширина ящика "старой" детали мм',
    'Box Height Before mm': 'Высота ящика "старой" детали мм',
    'Box Area Before m2': 'Площадь ящика "старой" детали мм2',
    'Box Volume Before m3': 'Объем ящика "старой" детали мм3',
    'Box Net Weight Before kg': 'Вес ящика "старой" детали кг',
    'Box Stacking Before units': 'Штабелирование ящика "старой" детали шт.',
    'Box Type After': 'Тип ящика "новой" детали',
    'Box Length After mm': 'Длина ящика "новой" детали мм',
    'Box Width After mm': 'Ширина ящика "новой" детали мм',
    'Box Height After mm': 'Высота ящика "новой" детали мм',
    'Box Area After m2': 'Площадь ящика "новой" детали мм2',
    'Box Volume After m3': 'Объем ящика "новой" детали мм3',
    'Box Net Weight After kg': 'Вес ящика "новой" детали кг',
    'Box Stacking After units': 'Штабелирование ящика "новой" детали шт.',
    'Boxes per Pallet Before units': 'Количество ящиков "старой" детали на 1 поддон шт.',
    'Boxes per Pallet After units': 'Количество ящиков "новой" детали на 1 поддон шт.',
    'Pallet Type Before': 'Тип поддона "старой" детали',
    'Pallet Length Before mm': 'Длина поддона "старой" детали мм',
    'Pallet Width Before mm': 'Ширина поддона "старой" детали мм',
    'Pallet Height Before mm': 'Высота поддона "старой" детали мм',
    'Pallet Area Before m2': 'Площадь поддона "старой" детали мм2',
    'Pallet Volume Before m3': 'Объем поддона "старой" детали мм3',
    'Pallet Gross Weight Before kg': 'Вес поддона "старой" детали кг',
    'Pallet Stacking Before units': 'Штабелирование поддона "старой" детали шт.',
    'Pallet Type After': 'Тип поддона "новой" детали',
    'Pallet Length After mm': 'Длина поддона "новой" детали мм',
    'Pallet Width After mm': 'Ширина поддона "новой" детали мм',
    'Pallet Height After mm': 'Высота поддона "новой" детали мм',
    'Pallet Area After m2': 'Площадь поддона "новой" детали мм2',
    'Pallet Volume After m3': 'Объем поддона "новой" детали мм3',
    'Pallet Gross Weight After kg': 'Вес поддона "новой" детали кг',
    'Pallet Stacking After units': 'Штабелирование поддона "новой" детали шт.',
    'Production Part Disposal': 'Использование "старых" деталей в производстве',
    'Interchangeable': 'Взаимозаменяемость "старых/новых" деталей',
    'Supplier Name Before': 'Название поставщика "старых" деталей',
    'Supplier Location Before': 'Страна поставщика "старых" деталей',
    'Supplier City Before': 'Город поставщика "старых" деталей',
    'Supplier Street Before': 'Улица поставщика "старых" деталей',
    'Supplier Building Before': 'Здание поставщика "старых" деталей',
    'Localization Before': 'Статус локализации поставщика "старых" деталей',
    'Supplier Name After': 'Название поставщика "новых" деталей',
    'Supplier Location After': 'Страна поставщика "новых" деталей',
    'Supplier City After': 'Город поставщика "новых" деталей',
    'Supplier Street After': 'Улица поставщика "новых" деталей',
    'Supplier Building After': 'Здание поставщика "новых" деталей',
    'Localization After': 'Статус локализации поставщика "новых" деталей',
    'Change Description (RUS)': 'Описание технического изменения',
    'Change Solution (RUS)': 'Решение технического изменения',
    'Color Code': 'Код цвета',
    'Color Name (RUS)': 'Название цвета',
    'Comments': 'Комментарии'
}

# 8. Порядок колонок
BP_DATA_COLUMNS_ORDER: List[str] = list(BP_DATA_TRANSLATION_COLS.keys())


# ========================================================================
# 9. РЕЗУЛЬТИРУЮЩИЕ КОЛОНКИ ШАГОВ ФУНКЦИИ process_bp_file()
# ========================================================================

# 1. Списки колонок предпоказа по порядку шагов
STEP_1_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'BOM Product', 'Change',
    'Update Type', 'Part No.',
    'Part Name(CHN)', 'In Stock'
]

STEP_2_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Batch Plan', 'Part No. Before',
    'Part No. After'
]

STEP_3_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'Batch Plan',
    'Part No. Before', 'Part No. After'
]

STEP_4_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'Batch Plan',
    'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Part No. After'
]

STEP_5_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status',
    'Batch Plan', 'Batch Fact',
    'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Part No. After'
]

STEP_6_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Part No. After'
]

STEP_7_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Quantity Parts in SS',
    'Part No. After'
]

STEP_8_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Part Name Before',
    'Part No. After', 'Part Name After'
]

STEP_9_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Supplier Name Before',
    'Part No. After', 'Supplier Name After'
]

STEP_10_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Part No. After',
    'Supplier Name Before', 'Localization Before',
    'Supplier Name After', 'Localization After'
]

STEP_11_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Localization Before',
    'Supplier Location Before', 'Supplier City Before',
    'Supplier Street Before', 'Supplier Building Before'
]

STEP_12_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'BOM Product Code', 'BOM Product Name',
    'Part No. After', 'Localization After',
    'Supplier Location After', 'Supplier City After',
    'Supplier Street After', 'Supplier Building After'
]

STEP_13_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'BOM Product Code', 'BOM Product Name',
    'Part No. Before', 'Part No. After',
    'Change Description'
]

STEP_14_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'Part No. Before', 'Part No. After',
    'Workcenter Code Before', 'Workcenter Name Before',
    'Workshop Code Before', 'Workshop Name Before',
    'Workcenter Code After', 'Workcenter Name After',
    'Workshop Code After', 'Workshop Name After'
]

STEP_15_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'Part No. Before', 'Part No. After',
    'Production Part Disposal'
]

STEP_16_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'Part No. Before', 'Part No. After',
    'Interchangeable'
]

STEP_17_FOCUS_COLUMNS: List[str] = CONF_TO_KEEP_COLS

STEP_18_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'Part No. Before', 'Part No. After',
    'Quantity Batches in SS'
]

STEP_19_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'Part No. Before', 'Part No. After',
    'Configuration for Old Parts Using Out',
    'Batches for Old Parts Using Out',
    'Transmission'
]

STEP_20_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'Part No. Before', 'Part Name Before',
    'Quantity per Box Before', 'Box Type Before',
    'Box Length Before mm', 'Box Width Before mm',
    'Box Height Before mm', 'Box Area Before m2',
    'Box Volume Before m3', 'Box Net Weight Before kg',
    'Box Stacking Before units', 'Boxes per Pallet Before units'
]

STEP_21_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
    'Part No. After', 'Part Name After',
    'Quantity per Box After', 'Box Type After',
    'Box Length After mm', 'Box Width After mm', 'Box Height After mm',
    'Pallet Length After mm', 'Pallet Width After mm', 'Pallet Height After mm',
    'Box Stacking Before units', 'Boxes per Pallet After units',
]

STEP_22_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
]

STEP_23_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
]

STEP_24_FOCUS_COLUMNS: List[str] = [
    'BP_No', 'Status', 'BOM Product',
    'Batch Plan', 'Batch Fact', 'Change Date',
]

# 2. Массив колонок предпоказа по порядку шагов
STEP_FOCUS_COLUMNS_MAP: List[List[str]] = [
    STEP_1_FOCUS_COLUMNS,
    STEP_2_FOCUS_COLUMNS,
    STEP_3_FOCUS_COLUMNS,
    STEP_4_FOCUS_COLUMNS,
    STEP_5_FOCUS_COLUMNS,
    STEP_6_FOCUS_COLUMNS,
    STEP_7_FOCUS_COLUMNS,
    STEP_8_FOCUS_COLUMNS,
    STEP_9_FOCUS_COLUMNS,
    STEP_10_FOCUS_COLUMNS,
    STEP_11_FOCUS_COLUMNS,
    STEP_12_FOCUS_COLUMNS,
    STEP_13_FOCUS_COLUMNS,
    STEP_14_FOCUS_COLUMNS,
    STEP_15_FOCUS_COLUMNS,
    STEP_16_FOCUS_COLUMNS,
    STEP_17_FOCUS_COLUMNS,
    STEP_18_FOCUS_COLUMNS,
    STEP_19_FOCUS_COLUMNS,
    STEP_20_FOCUS_COLUMNS,
    STEP_21_FOCUS_COLUMNS,
    STEP_22_FOCUS_COLUMNS,
    STEP_23_FOCUS_COLUMNS,
    STEP_24_FOCUS_COLUMNS,
]


# ====================================================================
# 10. НАСТРОЙКИ СТИЛЕЙ И ВИЗУАЛЬНОГО ОФОРМЛЕНИЯ EXCEL
# ====================================================================

# Цвета оформления (HEX-коды)
EXCEL_THEME_COLORS: Dict[str, str] = {
    'warehouse_bg': '#FDE9D9',      # Нежно-оранжевый для кладовщиков
    'warehouse_text': 'black',
    'header_bg': '#0F243E',         # Тёмно-синий для технических заголовков
    'header_text': 'white',
    'data_bg': '#FFFFFF'            # Белый фон для обычных ячеек
}

# Ширина колонок по их индексам
EXCEL_COLUMN_WIDTHS: Dict[int, int] = {
    0: 25, 1: 25, 2: 25, 3: 25, 4: 25, 5: 25, 6: 25, 7: 25,
    8: 80, 9: 25, 10: 25, 11: 25, 12: 25, 13: 25, 14: 25,
    15: 80, 16: 30, 17: 50, 18: 30, 19: 50,
    20: 30, 21: 30, 22: 30, 23: 30, 24: 30, 25: 30, 26: 30, 27: 30,
    28: 50, 29: 50, 30: 80, 31: 25, 32: 80, 33: 25,
    34: 100, 35: 100, 36: 25, 37: 25, 38: 100,
}


# ====================================================================
# 11. МАТРИЦА ЦЕХОВ / РАБОЧИХ СТАНЦИЙ
# ====================================================================

WORKSHOPS: Dict[str, Dict[str, Any]] = {
    'AS': {
        'name': 'Assembling',
        'prefixes': ['HA', 'HF'],
    },
    'CH': {
        'name': 'Chassis',
        'prefixes': ['HC'],
    },
    'COMP': {
        'name': 'Component',
        'prefixes': ['HN', 'HZ'],
    },
    'NOBO': {
        'name': 'Nobo',
        'prefixes': ['HL'],
    },
    'PAINT': {
        'name': 'Painting',
        'prefixes': ['HT'],
    },
    'SOFT': {
        'name': 'Software',
        'prefixes': ['EA'],
    },
    'STAMP': {
        'name': 'Stamping',
        'prefixes': ['HP', 'HE'],
    },
    'WELD': {
        'name': 'Welding',
        'prefixes': ['HW'],
    },
}

# Производный плоский индекс для быстрого lookup: префикс кода станции → код цеха
WORKSHOP_PREFIX_TO_CODE: Dict[str, str] = {
    prefix: code
    for code, info in WORKSHOPS.items()
    for prefix in info['prefixes']
}

# ====================================================================
# 12. МАТРИЦА КОД / ИМЯ МОДЕЛЕЙ
# ====================================================================

BOM_PRODUCT_MAP: Dict[str, str] = {
    'A01': 'Jolion',
    'A08': 'H3',
    'B02': 'F7',
    'B04': 'F7x',
    'B06': 'Dargo',
    'B16': 'H7',
}

# ====================================================================
# 13. СПРАВОЧНИКИ ДЛЯ ИНТЕРАКТИВНОГО ВВОДА
# ====================================================================

# Типы упаковки (для Box Type / Pallet Type)
PACKAGING_TYPE: List[str] = [
    'returnable',
    'non-returnable',
]

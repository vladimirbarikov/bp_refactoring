#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
# pylint: disable=import-outside-toplevel
"""
BP Refactoring Tool - Сопоставление и связывание технических изменений (Breakpoint).

Модуль является выделенным компонентом бизнес-логики пайплайна обработки BP.
Он отвечает за идентификацию, классификацию и сборку пар заменяемых деталей
(Before / After) на основе технологических признаков G-BOM.

Основная функция модуля:
    find_pairs(df_bp) -> List[Dict[str, Any]]

Алгоритм связывания пар (4 сценария):
    1. Жёсткое связывание Before и After по совпадению 'Part No.'
    2. Связывание по китайскому наименованию 'Part Name(CHN)'
       (применяется, когда номер детали изменился, а название осталось)
    3. Сборка непарных одиночных Before и After изменений
    4. Обработка логических операций из 'Update Type':
        • Add + Delete          → Delete = Before, Add = After
        • Add + Replace/Update  → Add = Before, Replace/Update = After

Экспортируемые функции:
    - find_pairs()                  — главная точка входа модуля
    - classify_row_before_after()   — классификация одной строки по колонке 'Change'
    - create_pair_dict()            — сборка матричной строки Before/After

Использование:
    from py_lib.pipeline.find_pairs import find_pairs
    pairs = find_pairs(df_bp)
    df_matrix = pd.DataFrame(pairs)

Версия: 1.0
Совместимость: Python 3.14.4+, Pandas 3.0.3+, OpenPyXL 3.1.5+
Поддержка: PLD Engineering Center
Дата создания: 2026-09-10
Лицензия: MIT
Статус: Production
"""
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd

from py_lib.engine.etl import str_convert, float_convert


def classify_row_before_after(
    row: pd.Series
) -> str:
    """
    Классифицирует отдельную строку файла изменений на Before, After или Unknown.

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
    change_val = str_convert(row.get('Change', ''))

    if change_val == 'Before Change':
        return 'Before'
    if change_val == 'After Change':
        return 'After'

    return 'Unknown'


def create_pair_dict(
    before_row: Optional[pd.Series],
    after_row: Optional[pd.Series]
) -> Dict[str, Any]:
    """
    Создаёт структурированный словарь пары деталей Before/After, полностью
    соответствующий целевой матричной спецификации колонок из core.py
    (BP_DATA_COLUMNS_ORDER).

    Функция собирает матричную строку из четырёх блоков:
        • Общие метаданные (BP_No, Status, Batch Plan, BOM Product Code, ...)
        • Параметры старой детали (Part No. Before, Supplier Name Before, ...)
        • Параметры новой детали (Part No. After, Supplier Name After, ...)
        • Инициализация будущих колонок (упаковка, локализация, конфигурация, ...)

    Если какая-то из сторон (Before/After) отсутствует (None), все её поля
    заполняются значением '-' (для строк) или np.nan (для чисел).

    Аргументы:
        before_row (Optional[pd.Series]): Строка с деталью Before или None.
        after_row (Optional[pd.Series]):  Строка с деталью After или None.

    Возвращается:
        Dict[str, Any]: Словарь с агрегированными данными матричной строки
                        для итоговой таблицы.
    """

    result: Dict[str, Any] = {}
    source_row = before_row if before_row is not None else after_row

    # === БЛОК 1: ОБЩИЕ МЕТАДАННЫЕ ДЛЯ ВСЕЙ СТРОКИ МАТРИЦЫ ===
    # Эти поля сквозные, они наследуются от любого присутствующего компонента в паре
    if source_row is not None:
        result['BP_No'] = str_convert(source_row.get('BP_No', ''))
        result['Status'] = str_convert(source_row.get('Status', '-'))
        result['Batch Plan'] = str_convert(source_row.get('In Stock', '-'))
        result['New Part Available Date'] = str_convert(source_row.get('New Part Available Date', '-'))
        result['BOM Product Code'] = str_convert(source_row.get('BOM Product', '-'))
        result['Change Description'] = str_convert(source_row.get('Change Description', '-'))
        result['Change Solution'] = str_convert(source_row.get('Solution', '-'))
        result['Color Code'] = str_convert(source_row.get('Color Code', '-'))
        result['Color Name'] = str_convert(source_row.get('Color Name', '-'))

        # Инициализируем будущие глобальные шаги ввода значениями по умолчанию
        result['Batch Fact'] = '-'
        result['Change Date'] = '-'
    else:
        for key in ['BP_No', 'Status', 'Batch Plan', 'New Part Available Date', 
                    'BOM Product Code', 'Change Description', 'Change Solution',
                    'Color Code', 'Color Name', 'Batch Fact', 'Change Date']:
            result[key] = '-'

    # === БЛОК 2: ПАРАМЕТРЫ СТАРЯ ДЕТАЛЬ "BEFORE" ===
    if before_row is not None:
        result['Part No. Before'] = str_convert(before_row.get('Part No.', '-'))
        result['Part Name Before'] = str_convert(before_row.get('Part Name(CHN)', '-'))
        result['Workcenter Code Before'] = str_convert(before_row.get('Workcenter No.', '-'))
        result['Workcenter Name Before'] = str_convert(before_row.get('Workcenter Name', '-'))
        result['Quantity per Vehicle Before'] = float_convert(before_row.get('Quantity', np.nan), default=np.nan)
        result['Supplier Name Before'] = str_convert(before_row.get('Supplier Name', '-'))
        result['Localization Before'] = str_convert(before_row.get('Localization', '-'))

        # Общие текстовые поля G-BOM подтягиваем из Before-строки, если они там есть
        result['Production Part Disposal'] = str_convert(before_row.get('Production Part Disposal', '-'))
        result['Interchangeable'] = str_convert(before_row.get('Interchangeable', '-'))
    else:
        result['Part No. Before'] = '-'
        result['Part Name Before'] = '-'
        result['Workcenter Code Before'] = '-'
        result['Workcenter Name Before'] = '-'
        result['Quantity per Vehicle Before'] = np.nan
        result['Supplier Name Before'] = '-'
        result['Localization Before'] = '-'
        result['Production Part Disposal'] = '-'
        result['Interchangeable'] = '-'

    # === БЛОК 3: ПАРАМЕТРЫ НОВАЯ ДЕТАЛЬ "AFTER" ===
    if after_row is not None:
        result['Part No. After'] = str_convert(after_row.get('Part No.', '-'))
        result['Part Name After'] = str_convert(after_row.get('Part Name(CHN)', '-'))
        result['Workcenter Code After'] = str_convert(after_row.get('Workcenter No.', '-'))
        result['Workcenter Name After'] = str_convert(after_row.get('Workcenter Name', '-'))
        result['Quantity per Vehicle After'] = float_convert(after_row.get('Quantity', np.nan), default=np.nan)
        result['Supplier Name After'] = str_convert(after_row.get('Supplier Name', '-'))
        result['Localization After'] = str_convert(after_row.get('Localization', '-'))

        # Если в Before-строке этих полей не было, забираем их из After-изменения
        if result['Production Part Disposal'] == '-':
            result['Production Part Disposal'] = str_convert(after_row.get('Production Part Disposal', '-'))
        if result['Interchangeable'] == '-':
            result['Interchangeable'] = str_convert(after_row.get('Interchangeable', '-'))
    else:
        result['Part No. After'] = '-'
        result['Part Name After'] = '-'
        result['Workcenter Code After'] = '-'
        result['Workcenter Name After'] = '-'
        result['Quantity per Vehicle After'] = np.nan
        result['Supplier Name After'] = '-'
        result['Localization After'] = '-'

    # === БЛОК 4: ИНИЦИАЛИЗАЦИЯ НОВЫХ СИСТЕМНЫХ И РАСЧЕТНЫХ КОЛОНОК ===
    # Будущие поля Parts
    result['Quantity Parts in SS'] = np.nan
    result['Part Net Weight Before kg'] = np.nan
    result['Part Net Weight After kg'] = np.nan

    # Будущие поля Batches
    result['Quantity Batches in SS'] = np.nan
    result['Batches for Old Parts Using Out'] = '-'

    # Будущие поля Configuration
    result['Configuration for Old Parts Using Out'] = '-'
    result['Transmission'] = '-'

    # Будущие поля Workshop
    result['Workshop Code Before'] = '-'
    result['Workshop Name Before'] = '-'
    result['Workshop Code After'] = '-'
    result['Workshop Name After'] = '-'

    # Будущие поля Packaging Before
    result['Quantity per Box Before'] = np.nan
    result['Box Type Before'] = '-'
    result['Box Length Before mm'] = np.nan
    result['Box Width Before mm'] = np.nan
    result['Box Height Before mm'] = np.nan
    result['Box Area Before m2'] = np.nan
    result['Box Volume Before m3'] = np.nan
    result['Box Net Weight Before kg'] = np.nan
    result['Boxes per Pallet Before units'] = np.nan
    result['Box Stacking Before units'] = np.nan
    result['Pallet Type Before'] = '-'
    result['Pallet Length Before mm'] = np.nan
    result['Pallet Width Before mm'] = np.nan
    result['Pallet Height Before mm'] = np.nan
    result['Pallet Area Before m2'] = np.nan
    result['Pallet Volume Before m3'] = np.nan
    result['Pallet Gross Weight Before kg'] = np.nan
    result['Pallet Stacking Before units'] = np.nan

    # Будущие поля Packaging After
    result['Quantity per Box After'] = np.nan
    result['Box Type After'] = '-'
    result['Box Length After mm'] = np.nan
    result['Box Width After mm'] = np.nan
    result['Box Height After mm'] = np.nan
    result['Box Area After m2'] = np.nan
    result['Box Volume After m3'] = np.nan
    result['Box Net Weight After kg'] = np.nan
    result['Box Stacking After units'] = np.nan
    result['Boxes per Pallet After units'] = np.nan
    result['Pallet Type After'] = '-'
    result['Pallet Length After mm'] = np.nan
    result['Pallet Width After mm'] = np.nan
    result['Pallet Height After mm'] = np.nan
    result['Pallet Area After m2'] = np.nan
    result['Pallet Volume After m3'] = np.nan
    result['Pallet Gross Weight After kg'] = np.nan
    result['Pallet Stacking After units'] = np.nan

    # Будущие поля Supplier Before
    result['Supplier Location Before'] = '-'
    result['Supplier City Before'] = '-'
    result['Supplier Street Before'] = '-'
    result['Supplier Building Before'] = '-'

    # Будущие поля Supplier After
    result['Supplier Location After'] = '-'
    result['Supplier City After'] = '-'
    result['Supplier Street After'] = '-'
    result['Supplier Building After'] = '-'

    # Остальные будущие поля
    result['BOM Product Name'] = '-'
    result['Comments'] = '-'

    return result


def find_pairs(
    df_bp: pd.DataFrame
) -> List[Dict[str, Any]]:
    """
    Выполняет сквозной анализ плоского датафрейма G-BOM и связывает
    компоненты в логические пары в соответствии с инженерной матрицей
    колонок BP_DATA_COLUMNS_ORDER.

    Алгоритм работы (4 последовательных сценария):

        СЦЕНАРИЙ 1: Жёсткое связывание Before / After по 'Part No.'
            Классификация строк по колонке 'Change':
                'Before Change' → Before
                'After Change'  → After

        СЦЕНАРИЙ 2: Связывание по китайскому наименованию 'Part Name(CHN)'
            Применяется для модификаций, когда номер детали изменился,
            а название осталось прежним.

        СЦЕНАРИЙ 3: Сборка непарных одиночных Before и After изменений
            Все оставшиеся после Сценариев 1–2 строки добавляются как
            одиночные матричные записи (Before без After или наоборот).

        СЦЕНАРИЙ 4: Обработка логических операций из 'Update Type'
            Запускается для строк, у которых колонка 'Change' пустая:
                • Add + Delete          → Delete = Before, Add = After
                • Add + Replace/Update  → Add = Before, Replace/Update = After
                • Оставшиеся Delete / Add / Replace-Update → одиночные строки

    Аргументы:
        df_bp (pd.DataFrame): Исходный очищенный DataFrame с данными BP файла.

    Возвращается:
        List[Dict[str, Any]]: Список словарей, где каждый словарь представляет
            собой готовую матричную строку Before/After для итогового отчёта.
    """
    df_working = df_bp.copy()
    # Классифицируем строки на Before, After или Unknown
    df_working['_direction'] = df_working.apply(classify_row_before_after, axis=1)

    before_rows = df_working[df_working['_direction'] == 'Before'].copy()
    after_rows = df_working[df_working['_direction'] == 'After'].copy()
    unknown_rows = df_working[df_working['_direction'] == 'Unknown'].copy()

    pairs: List[Dict[str, Any]] = []
    used_before = set()
    used_after = set()

    # === СЦЕНАРИЙ 1: Жесткое связывание Before и After по 'Part No.' ===
    for before_idx, before_row in before_rows.iterrows():
        before_part_no = str_convert(before_row.get('Part No.', ''))
        if before_part_no in ('', '-'):
            continue

        for after_idx, after_row in after_rows.iterrows():
            if after_idx in used_after:
                continue
            if str_convert(after_row.get('Part No.', '')) == before_part_no:
                pairs.append(create_pair_dict(before_row, after_row))
                used_before.add(before_idx)
                used_after.add(after_idx)
                break

    # === СЦЕНАРИЙ 2: Связывание по китайскому наименованию 'Part Name(CHN)' ===
    # Применяется для модификаций, когда название детали одинаковое, а номер изменился
    for before_idx, before_row in before_rows.iterrows():
        if before_idx in used_before:
            continue
        before_part_name = str_convert(before_row.get('Part Name(CHN)', ''))
        if before_part_name in ('', '-'):
            continue

        for after_idx, after_row in after_rows.iterrows():
            if after_idx in used_after:
                continue
            if str_convert(after_row.get('Part Name(CHN)', '')) == before_part_name:
                pairs.append(create_pair_dict(before_row, after_row))
                used_before.add(before_idx)
                used_after.add(after_idx)
                break

    # === СЦЕНАРИЙ 3: Сборка непарных одиночных Before и After изменений ===
    for before_idx, before_row in before_rows.iterrows():
        if before_idx not in used_before:
            pairs.append(create_pair_dict(before_row, None))

    for after_idx, after_row in after_rows.iterrows():
        if after_idx not in used_after:
            pairs.append(create_pair_dict(None, after_row))

    # === СЦЕНАРИЙ 4: Обработка логических операций из 'Update Type' ===
    # Запускается для строк, у которых колонка 'Change' отсутствовала или была пустой
    if not unknown_rows.empty:
        delete_rows = []
        add_rows = []
        replace_update_rows = []

        for _, row in unknown_rows.iterrows():
            u_type = str_convert(row.get('Update Type', ''))
            if u_type == 'Delete':
                delete_rows.append(row)
            elif u_type == 'Add':
                add_rows.append(row)
            elif u_type in ('Replace', 'Update'):
                replace_update_rows.append(row)

        # 4.1 Связывание Add + Delete (Замена старой детали на точно такую же новую)
        used_delete = set()
        used_add_delete = set()

        for add_idx, add_row in enumerate(add_rows):
            add_part_no = str_convert(add_row.get('Part No.', ''))
            if add_part_no in ('', '-'):
                continue
            for del_idx, del_row in enumerate(delete_rows):
                if del_idx in used_delete:
                    continue
                if str_convert(del_row.get('Part No.', '')) == add_part_no:
                    pairs.append(create_pair_dict(del_row, add_row))
                    used_delete.add(del_idx)
                    used_add_delete.add(add_idx)
                    break

        # 4.2 Связывание Add + Replace/Update (Связывание по номеру детали)
        used_replace_update = set()
        used_add_replace = set()

        for add_idx, add_row in enumerate(add_rows):
            if add_idx in used_add_delete or add_idx in used_add_replace:
                continue
            add_part_no = str_convert(add_row.get('Part No.', ''))
            if add_part_no in ('', '-'):
                continue
            for ru_idx, ru_row in enumerate(replace_update_rows):
                if ru_idx in used_replace_update:
                    continue
                if str_convert(ru_row.get('Part No.', '')) == add_part_no:
                    pairs.append(create_pair_dict(add_row, ru_row))
                    used_add_replace.add(add_idx)
                    used_replace_update.add(ru_idx)
                    break

        # 4.3 Фиксация оставшихся изолированных операций как одиночных матричных строк
        for del_idx, del_row in enumerate(delete_rows):
            if del_idx not in used_delete:
                pairs.append(create_pair_dict(del_row, None))

        for add_idx, add_row in enumerate(add_rows):
            if add_idx not in used_add_delete and add_idx not in used_add_replace:
                pairs.append(create_pair_dict(None, add_row))

        for ru_idx, ru_row in enumerate(replace_update_rows):
            if ru_idx not in used_replace_update:
                pairs.append(create_pair_dict(None, ru_row))

    return pairs

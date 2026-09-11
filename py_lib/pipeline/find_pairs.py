#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
"""
Модуль сопоставления и связывания инженерных изменений (Change Matching Layer).

Спроектирован как выделенный компонент бизнес-логики для идентификации,
классификации и сборки пар заменяемых деталей (Before / After) на основе 
технологических признаков G-BOM.

Версия: 1.0
Совместимость: Python 3.14.4, Pandas 3.0.3, OpenPyXL 3.1.5
Поддержка: PLD Engineering Center
Дата создания: 2026-09-10
Лицензия: MIT
Статус: Production
"""
from typing import Any, Dict, List, Optional

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
    Создаёт структурированный словарь пары деталей Before/After для сквозного учета.

    Аргументы:
        before_row (Optional[pd.Series]): Строка с деталью Before или None.
        after_row (Optional[pd.Series]): Строка с деталью After или None.
    
    Возвращается:
        Dict[str, Any]: Словарь с агрегированными данными для итоговой таблицы.
    """
    result: Dict[str, Any] = {}
    source_row = before_row if before_row is not None else after_row

    if source_row is not None:
        result['BP_No'] = str_convert(source_row.get('BP_No', ''))
        result['Batch plan'] = str_convert(source_row.get('In Stock', ''))
        result['New Part Available Date'] = str_convert(source_row.get('New Part Available Date', ''))
        result['BOM Product'] = str_convert(source_row.get('BOM Product', ''))
        result['Production Part Disposal'] = str_convert(source_row.get('Production Part Disposal', ''))
        result['Interchangeable'] = str_convert(source_row.get('Interchangeable', ''))
        result['Change Description'] = str_convert(source_row.get('Change Description (RUS)', ''))
        result['Change Solution'] = str_convert(source_row.get('Solution (RUS)', ''))
        result['Color Code'] = str_convert(source_row.get('Color Code', ''))
        result['Color Name (RUS)'] = str_convert(source_row.get('Color Name (RUS)', ''))
        result['Status'] = str_convert(source_row.get('Status', ''))
        result['Quantity in SS'] = float_convert(source_row.get('Quantity in SS', 0))
    else:
        for key in ['BP_No', 'Batch plan', 'New Part Available Date', 'BOM Product',
                    'Production Part Disposal', 'Interchangeable', 'Change Description', 
                    'Change Solution', 'Color Code', 'Color Name (RUS)', 'Status']:
            result[key] = ''
        result['Quantity in SS'] = 0.0

    # Заполнение ветки Before
    if before_row is not None:
        result['Part No. Before'] = str_convert(before_row.get('Part No.', ''))
        result['Part Name Before'] = str_convert(before_row.get('Part Name (RUS)', ''))
        result['Quantity per Vehicle Before'] = float_convert(before_row.get('Quantity', 0))
        result['Workcenter No. Before'] = str_convert(before_row.get('Workcenter No.', ''))
        result['Workcenter Name Before'] = str_convert(before_row.get('Workcenter Name', ''))
        result['Supplier Name Before'] = str_convert(before_row.get('Supplier Name (RUS)', ''))
        result['Localization Before'] = str_convert(before_row.get('Localization', ''))
    else:
        result['Part No. Before'] = ''
        result['Part Name Before'] = ''
        result['Quantity per Vehicle Before'] = 0.0
        for key in ['Workcenter No. Before', 'Workcenter Name Before', 'Supplier Name Before', 'Localization Before']:
            result[key] = ''

    # Заполнение ветки After
    if after_row is not None:
        result['Part No. After'] = str_convert(after_row.get('Part No.', ''))
        result['Part Name After'] = str_convert(after_row.get('Part Name (RUS)', ''))
        result['Quantity per Vehicle After'] = float_convert(after_row.get('Quantity', 0))
        result['Workcenter No. After'] = str_convert(after_row.get('Workcenter No.', ''))
        result['Workcenter Name After'] = str_convert(after_row.get('Workcenter Name', ''))
        result['Supplier Name After'] = str_convert(after_row.get('Supplier Name (RUS)', ''))
        result['Localization After'] = str_convert(after_row.get('Localization', ''))
    else:
        result['Part No. After'] = ''
        result['Part Name After'] = ''
        result['Quantity per Vehicle After'] = 0.0
        for key in ['Workcenter No. After', 'Workcenter Name After', 'Supplier Name After', 'Localization After']:
            result[key] = ''

    # Инициализация полей под будущую обработку конвейером
    result['Batch fact'] = ''
    result['Change Date'] = None
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
) -> List[Dict[str, Any]]:
    """
    Выполняет сквозной анализ плоского датафрейма и связывает компоненты в логические пары.

    Алгоритм работы:
        1. Классифицирует строки по колонке 'Change' (Before Change/After Change)
            и по колонке 'Update Type' (Delete, Add, Replace, Update)
        2. Связывает явные Before/After по Part No. или Part Name
        3. Обрабатывает комбинации Add + Delete (Add=After, Delete=Before)
        4. Обрабатывает комбинации Add + Replace/Update (Add=Before, Replace/Update=After)
        5. Оставшиеся непарные строки добавляются как одиночные
    
    Аргументы:
        df_bp (pd.DataFrame): Исходный DataFrame с данными BP файла.

    Возвращается:
        List[Dict[str, Any]]: Список словарей, каждый словарь представляет одну пару
            или одиночную деталь с заполненными полями Before/After.
    """
    df_working = df_bp.copy()
    df_working['_direction'] = df_working.apply(classify_row_before_after, axis=1)

    before_rows = df_working[df_working['_direction'] == 'Before'].copy()
    after_rows = df_working[df_working['_direction'] == 'After'].copy()
    unknown_rows = df_working[df_working['_direction'] == 'Unknown'].copy()

    pairs: List[Dict[str, Any]] = []
    used_before = set()
    used_after = set()

    # Связывание по Part No.
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

    # Связывание по Part Name (RUS)
    for before_idx, before_row in before_rows.iterrows():
        if before_idx in used_before:
            continue
        before_part_name = str_convert(before_row.get('Part Name (RUS)', ''))
        if before_part_name in ('', '-'):
            continue

        for after_idx, after_row in after_rows.iterrows():
            if after_idx in used_after:
                continue
            if str_convert(after_row.get('Part Name (RUS)', '')) == before_part_name:
                pairs.append(create_pair_dict(before_row, after_row))
                used_before.add(before_idx)
                used_after.add(after_idx)
                break

    # Одиночные Before и After
    for before_idx, before_row in before_rows.iterrows():
        if before_idx not in used_before:
            pairs.append(create_pair_dict(before_row, None))

    for after_idx, after_row in after_rows.iterrows():
        if after_idx not in used_after:
            pairs.append(create_pair_dict(None, after_row))

    # Обработка операций из Update Type (Delete, Add, Replace, Update)
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

        # Связывание Add + Delete
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

        # Связывание Add + Replace/Update
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

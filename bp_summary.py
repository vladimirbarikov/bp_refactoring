#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль bp_summary последовательно преобразует обработанные BP DataFrame в итоговый df_summary_breakpoint
"""

import os
import sys
import io
from typing import Dict, List, Optional

import pandas as pd

# Путь к сетевым папкам с файлами партий (подсказка для пользователя)
BATCH_FILES_PATH = r'\\hmmr_share\LD\Custom Clearance\Поставки\Серийный KD parts'

# Для Windows консоли
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


def clear_screen():
    """Очистка экрана консоли"""
    os.system('cls' if os.name == 'nt' else 'clear')


def wait_for_user(prompt="\nНажмите Enter для продолжения..."):
    """Ожидание нажатия Enter"""
    try:
        input(prompt)
    except KeyboardInterrupt:
        print("\n\nПрограмма прервана пользователем (Ctrl+C)")
        sys.exit(0)
    except EOFError:
        print("\n\nОбнаружен конец ввода. Программа завершена.")
        sys.exit(0)


def print_step_header(step_num, total_steps, description):
    """Вывод заголовка шага"""
    print("\n" + "=" * 60)
    print(f"ШАГ {step_num}/{total_steps}: {description}")
    print("=" * 60)


def save_state_before_step(df):
    """
    Сохраняет состояние DataFrame перед выполнением шага
    Возвращает сохранённую копию DataFrame
    """
    if df is not None:
        print("  [Сохранено состояние перед шагом]")
        return df.copy(deep=True)
    return None


def restore_state(saved_df, step_name):
    """
    Восстанавливает сохранённое состояние DataFrame
    Возвращает восстановленный DataFrame
    """
    if saved_df is not None:
        print(f"  [Восстанавливаем состояние перед шагом: {step_name}]")
        return saved_df.copy(deep=True)
    print("  Ошибка: нет сохранённого состояния для восстановления!")
    return None


def confirm_step(step_name, df_current, saved_state):
    """
    Запрашивает у пользователя подтверждение после выполнения шага
    Возвращает (continue_flag, df, new_saved_state)
    """
    print(f"\n  Шаг '{step_name}' выполнен.")
    try:
        user_input = input(
            "\nПроверьте результат. Если всё корректно, нажмите Enter. Если нужно повторить шаг, введите 'retry': "
        ).strip().lower()
    except KeyboardInterrupt:
        print("\n\nПрограмма прервана пользователем (Ctrl+C)")
        sys.exit(0)
    except EOFError:
        print("\n\nОбнаружен конец ввода. Программа завершена.")
        sys.exit(0)

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


def safe_float_convert(value, default=0.0) -> float:
    """
    Безопасное преобразование значения в float
    """
    if value is None or value == '' or value == '-':
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def safe_str_convert(value, default='') -> str:
    """
    Безопасное преобразование значения в строку
    """
    if value is None:
        return default
    try:
        return str(value).strip()
    except (ValueError, TypeError):
        return default


def load_configuration_file(config_filename: str = 'configuration.xlsx') -> Optional[pd.DataFrame]:
    """Загрузка конфигурационного файла (лист 'common')"""
    if not os.path.exists(config_filename):
        print(f"Предупреждение: Файл конфигурации '{config_filename}' не найден")
        print("Некоторые поля будут пустыми")
        return None

    try:
        # Загружаем лист 'common'
        df_config = pd.read_excel(config_filename, sheet_name='common')
        print(f"Файл конфигурации '{config_filename}' (лист 'common') загружен: {df_config.shape[0]} строк")

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


def show_dataframe_preview(df, step_name, max_rows=5, focus_columns=None, max_colwidth=40):
    """
    Отображает первые строки DataFrame для визуального контроля
    Числовые поля отображаются как числа, 0.0 заменяется на '-' для удобства чтения
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

    # Создаем копию DataFrame для отображения
    preview_df = df[display_cols].head(max_rows).copy()

    # Обрабатываем каждую колонку в зависимости от типа данных
    for col in preview_df.columns:
        # Определяем тип данных в колонке
        if pd.api.types.is_float_dtype(preview_df[col]) or pd.api.types.is_integer_dtype(preview_df[col]):
            # Числовые поля: заменяем 0.0 и NaN на '-', остальные числа оставляем как есть
            preview_df[col] = preview_df[col].fillna(0.0)
            preview_df[col] = preview_df[col].apply(
                lambda x: '-' if x == 0.0 else (str(x) if isinstance(x, (int, float)) else x)
            )
        else:
            # Строковые поля: заполняем NaN и обрезаем длинные значения
            preview_df[col] = preview_df[col].fillna('-').astype(str)
            preview_df[col] = preview_df[col].apply(
                lambda x: (x[:max_colwidth] + '…') if len(x) > max_colwidth else x
            )

    # Выводим с помощью pandas
    with pd.option_context(
        'display.max_columns', len(display_cols),
        'display.width', None,
        'display.max_colwidth', max_colwidth,
        'display.show_dimensions', False,
        'display.unicode.east_asian_width', True
    ):
        print(preview_df.to_string(index=False))
    print()


def classify_row_before_after(row: pd.Series) -> str:
    """
    Классификация детали на 'Before', 'After' или 'Unknown'

    Приоритет 1: колонка 'Change'
        - 'Before Change' → Before
        - 'After Change' → After

    Приоритет 2: если Change пустой, то:
        - Строки с 'Delete' → 'Unknown' (для обработки в паре с Add)
        - Строки с 'Add' → 'Unknown' (для обработки в паре с Delete/Replace/Update)
        - Строки с 'Replace' → 'Unknown' (для обработки в паре с Add)
        - Строки с 'Update' → 'Unknown' (для обработки в паре с Add)

    ВСЕ строки без явного Change попадают в Unknown, чтобы связывание 
    происходило только внутри find_pairs по правилам:
    - Add + Delete → Add = After, Delete = Before
    - Add + Replace → Add = Before, Replace = After
    - Add + Update → Add = Before, Update = After
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


def create_pair_dict(before_row: Optional[pd.Series], after_row: Optional[pd.Series]) -> Dict:
    """Создает словарь для одной строки итоговой таблицы из пары Before/After"""
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
    result['Batches for old parts using out'] = 0.0
    result['Transmission'] = ''
    result['Quantity per Box Before'] = 0.0
    result['Quantity per Box After'] = 0.0
    result['Box Before (L-W-H) mm'] = ''
    result['Box After (L-W-H) mm'] = ''
    result['Pallet Before (L-W-H) mm'] = ''
    result['Pallet After (L-W-H) mm'] = ''
    result['Comments'] = ''

    return result


def find_pairs(df_bp: pd.DataFrame) -> List[Dict]:
    """
    Поиск пар Before/After деталей

    Алгоритм:
    1. Классифицируем строки через classify_row_before_after()
    2. Разделяем на группы: Before, After, Unknown
    3. Before и After связываем в пары по Part No. или Part Name (RUS)
    4. Unknown обрабатываем по правилам Add+Delete, Add+Replace, Add+Update
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


def user_input_for_single_bp(df_current: pd.DataFrame, bp_number: str) -> pd.DataFrame:
    """
    Интерактивный ввод Batch fact и Change Date для одного BP файла
    Данные вводятся один раз для всего BP, а не для каждой строки
    """
    print(f"\n--- Ввод данных для BP {bp_number} ---")
    print(f"Всего строк для обработки: {len(df_current)}")

    df_result = df_current.copy()

    print("\nВведите общие данные для всего технического изменения:")

    # Batch fact
    current_batch_fact = safe_str_convert(df_result.iloc[0].get('Batch fact', '')) if len(df_result) > 0 else ''
    print(f"Текущее Batch fact: {current_batch_fact if current_batch_fact else '(пусто)'}")
    batch_fact_input = input("Введите Batch fact (или Enter, чтобы оставить пустым): ").strip()

    # Change Date
    current_change_date = safe_str_convert(df_result.iloc[0].get('Change Date', '')) if len(df_result) > 0 else ''
    print(f"Текущее Change Date: {current_change_date if current_change_date else '(пусто)'}")
    change_date_input = input("Введите Change Date (ГГГГ-ММ-ДД или Enter, чтобы оставить пустым): ").strip()

    if batch_fact_input:
        df_result['Batch fact'] = batch_fact_input
        print(f"  Batch fact '{batch_fact_input}' применён ко всем {len(df_result)} строкам")

    if change_date_input:
        df_result['Change Date'] = change_date_input
        print(f"  Change Date '{change_date_input}' применён ко всем {len(df_result)} строкам")

    if not batch_fact_input and not change_date_input:
        print("  Данные не введены. Будут заполнены позже в Excel.")

    return df_result


def config_lookup_for_single_bp(df_current: pd.DataFrame, df_config: Optional[pd.DataFrame]) -> pd.DataFrame:
    """
    Поиск значений в конфигурационном файле для одного BP файла
    """
    if df_config is None or df_config.empty:
        print("  Конфигурационный файл не загружен. Пропускаем поиск.")
        return df_current

    df_result = df_current.copy()

    if 'BOM Product' not in df_config.columns or 'Quantity vehicle in batch' not in df_config.columns:
        print("  Ошибка: В конфигурационном файле отсутствуют обязательные колонки")
        return df_result

    df_config['BOM Product'] = df_config['BOM Product'].astype(str).str.strip()
    if 'Batch code' in df_config.columns:
        df_config['Batch code'] = df_config['Batch code'].astype(str).str.strip()
    if 'Configuration' in df_config.columns:
        df_config['Configuration'] = df_config['Configuration'].astype(str).str.strip()
    if 'Transmission' in df_config.columns:
        df_config['Transmission'] = df_config['Transmission'].astype(str).str.strip()

    for idx, row in df_result.iterrows():
        bom_product = safe_str_convert(row.get('BOM Product', ''))
        batch_fact = safe_str_convert(row.get('Batch fact', ''))

        if bom_product == '' or bom_product == '-':
            continue

        config_matches = df_config[df_config['BOM Product'] == bom_product]
        if config_matches.empty:
            continue

        quantity_vehicle_in_batch = config_matches.iloc[0].get('Quantity vehicle in batch')
        quantity_in_ss = row.get('Quantity in SS', 0)

        qty_in_ss = safe_float_convert(quantity_in_ss, 0.0)
        qty_vehicle = safe_float_convert(quantity_vehicle_in_batch, 1.0)

        if qty_vehicle > 0:
            qty_batches = round(qty_in_ss / qty_vehicle, 2)
            df_result.at[idx, 'Quantity batches in SS'] = qty_batches
            print(f"  {bom_product}: Quantity batches in SS = {qty_batches}")

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


def load_batch_file(batch_name: str, search_path: str = BATCH_FILES_PATH) -> Optional[pd.DataFrame]:
    """Загрузка Excel файла партии (заглушка, будет реализована позже)"""
    batch_name = safe_str_convert(batch_name)
    if not batch_name or batch_name == '' or batch_name == '-':
        return None

    if len(batch_name) >= 4:
        file_number = batch_name[3:]
        file_prefix = batch_name[:3]
        filename = f"{file_number} {file_prefix}.xlsx"
    else:
        filename = f"{batch_name}.xlsx"

    full_path = os.path.join(search_path, filename)

    if not os.path.exists(full_path):
        print(f"  Файл не найден: {full_path}")
        return None

    try:
        df_batch = pd.read_excel(full_path)
        print(f"  Файл загружен: {filename}")
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


def batch_file_loader_for_single_bp(df_current: pd.DataFrame) -> pd.DataFrame:
    """
    Загрузка файлов партий для одного BP файла (заглушка)
    TODO: Реализовать после получения образцов файлов
    """
    print("\n  Проверка файлов партий...")
    print("  ВНИМАНИЕ: Функционал в разработке. Поля останутся пустыми.")

    df_result = df_current.copy()

    for _, row in df_result.iterrows():
        batch_fact = safe_str_convert(row.get('Batch fact', ''))
        if batch_fact and batch_fact != '' and batch_fact != '-':
            print(f"    Проверка партии: {batch_fact}")
            df_batch = load_batch_file(batch_fact)
            if df_batch is not None:
                print("      Файл загружен (TODO: извлечение данных)")

    return df_result


def process_single_bp(
    bp_number: str,
    df_bp: pd.DataFrame,
    df_config: Optional[pd.DataFrame],
    interactive: bool = True
) -> pd.DataFrame:
    """
    Обработка одного BP файла с подтверждением каждого шага:
    1. Поиск пар Before/After
    2. Пользовательский ввод
    3. Поиск в конфигурации
    4. Загрузка файлов партий
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
        print_step_header(1, 4, "Поиск пар Before/After")
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
            print_step_header(2, 4, "Ввод данных пользователем")
            df_current = user_input_for_single_bp(df_current, bp_number)
            show_dataframe_preview(
                df_current, "Ввод данных пользователем",
                focus_columns=['BP_No', 'Batch fact', 'Change Date', 'Part No. Before', 'Part No. After']
            )

            continue_flag, df_current, saved_state = confirm_step("Ввод данных пользователем", df_current, saved_state)
            if continue_flag:
                break

    # Шаг 3: Поиск в конфигурации
    while True:
        print_step_header(3, 4, "Поиск в конфигурационном файле")
        df_current = config_lookup_for_single_bp(df_current, df_config)
        show_dataframe_preview(
            df_current, "Поиск в конфигурации",
            focus_columns=['BP_No', 'BOM Product', 'Quantity in SS', 'Quantity batches in SS', 'Configuration for old parts using out']
        )

        continue_flag, df_current, saved_state = confirm_step("Поиск в конфигурационном файле", df_current, saved_state)
        if continue_flag:
            break

    # Шаг 4: Загрузка файлов партий
    while True:
        print_step_header(4, 4, "Проверка файлов партий")
        df_current = batch_file_loader_for_single_bp(df_current)
        show_dataframe_preview(
            df_current, "Проверка файлов партий",
            focus_columns=['BP_No', 'Batch fact', 'Quantity per Box Before', 'Quantity per Box After']
        )

        continue_flag, df_current, saved_state = confirm_step("Проверка файлов партий", df_current, saved_state)
        if continue_flag:
            break

    print(f"\n  BP {bp_number} обработан. Добавлено строк: {len(df_current)}")

    return df_current


def main(processed_results: Dict[str, pd.DataFrame], interactive: bool = True) -> pd.DataFrame:
    """
    Главная функция модуля summary_breakpoint_table
    Последовательная обработка каждого BP файла
    """
    clear_screen()

    print("""
        ╔══════════════════════════════════════════════════════════════╗
        ║           SUMMARY BREAKPOINT TABLE GENERATOR v1.0            ║
        ║        Формирование итоговой таблицы из BP файлов            ║
        ╚══════════════════════════════════════════════════════════════╝
        """)

    print(f"Найдено BP файлов для обработки: {len(processed_results)}")

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
            df_summary[col] = ''

    df_summary = df_summary[column_order]

    # Итоги
    clear_screen()
    print("\n" + "=" * 60)
    print("ИТОГОВАЯ ТАБЛИЦА SUMMARY BREAKPOINT")
    print("=" * 60)
    print(f"Обработано BP файлов: {len(all_processed_dfs)}/{len(processed_results)}")
    print(f"Всего строк в итоговой таблице: {len(df_summary)}")
    print(f"Колонок: {len(df_summary.columns)}")

    return df_summary


if __name__ == "__main__":
    print("Этот модуль предназначен для импорта из bp_main")
    print("Использование: from bp_summary import main")

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль summary_breakpoint_table
Преобразует обработанные BP DataFrame в итоговый df_summary_breakpoint
Последовательная обработка каждого BP файла
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
    """Загрузка конфигурационного файла"""
    if not os.path.exists(config_filename):
        print(f"Предупреждение: Файл конфигурации '{config_filename}' не найден")
        print("Некоторые поля будут пустыми")
        return None

    try:
        df_config = pd.read_excel(config_filename)
        print(f"Файл конфигурации '{config_filename}' загружен: {df_config.shape[0]} строк")
        return df_config
    except FileNotFoundError:
        print(f"Ошибка: Файл конфигурации '{config_filename}' не найден")
        return None
    except PermissionError:
        print(f"Ошибка: Нет прав для чтения файла '{config_filename}'")
        return None
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
        print(f"Ошибка: Файл '{config_filename}' повреждён: {e}")
        return None
    except Exception as e:
        print(f"Непредвиденная ошибка при загрузке: {e}")
        return None


def classify_row_before_after(row: pd.Series) -> str:
    """
    Шаг 1: Классификация детали на 'Before' или 'After'
    Правила:
    - Delete → всегда Before
    - Add, Replace, Update → требуют анализа пар (Unknown)
    """
    change_val = safe_str_convert(row.get('Change', ''))
    if change_val == 'Before Change':
        return 'Before'
    elif change_val == 'After Change':
        return 'After'

    update_type = safe_str_convert(row.get('Update Type', ''))

    if update_type == 'Delete':
        return 'Before'
    elif update_type in ['Add', 'Replace', 'Update']:
        return 'Unknown'
    else:
        return 'Unknown'


def create_pair_dict(before_row: Optional[pd.Series], after_row: Optional[pd.Series]) -> Dict:
    """Создает словарь для одной строки итоговой таблицы из пары Before/After"""
    result = {}

    source_row = before_row if before_row is not None else after_row
    if source_row is not None:
        result['BP_No'] = source_row.get('BP_No', '')
        result['Batch plan'] = source_row.get('In Stock', '')
        result['New Part Available Date'] = source_row.get('New Part Available Date', '')
        result['BOM Product'] = source_row.get('BOM Product', '')
        result['Production Part Disposal'] = source_row.get('Production Part Disposal', '')
        result['Interchangeable'] = source_row.get('Interchangeable', '')
        result['Change Description'] = source_row.get('Change Description (RUS)', '')
        result['Change Solution'] = source_row.get('Solution (RUS)', '')
        result['Color Code'] = source_row.get('Color Code', '')
        result['Color Name (RUS)'] = source_row.get('Color Name (RUS)', '')
        result['Status'] = source_row.get('Status', '')
        result['Quantity in SS'] = source_row.get('Quantity in SS', '')
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
        result['Quantity in SS'] = ''

    # Before деталь
    if before_row is not None:
        result['Part No. Before'] = before_row.get('Part No.', '')
        result['Part Name Before'] = before_row.get('Part Name (RUS)', '')
        result['Quantity per Vehicle Before'] = before_row.get('Quantity', '')
        result['Workcenter No. Before'] = before_row.get('Workcenter No.', '')
        result['Workcenter Name Before'] = before_row.get('Workcenter Name', '')
        result['Supplier Name Before'] = before_row.get('Supplier Name (RUS)', '')
        result['Localization Before'] = before_row.get('Localization', '')
    else:
        result['Part No. Before'] = '-'
        result['Part Name Before'] = '-'
        result['Quantity per Vehicle Before'] = '-'
        result['Workcenter No. Before'] = '-'
        result['Workcenter Name Before'] = '-'
        result['Supplier Name Before'] = '-'
        result['Localization Before'] = '-'

    # After деталь
    if after_row is not None:
        result['Part No. After'] = after_row.get('Part No.', '')
        result['Part Name After'] = after_row.get('Part Name (RUS)', '')
        result['Quantity per Vehicle After'] = after_row.get('Quantity', '')
        result['Workcenter No. After'] = after_row.get('Workcenter No.', '')
        result['Workcenter Name After'] = after_row.get('Workcenter Name', '')
        result['Supplier Name After'] = after_row.get('Supplier Name (RUS)', '')
        result['Localization After'] = after_row.get('Localization', '')
    else:
        result['Part No. After'] = '-'
        result['Part Name After'] = '-'
        result['Quantity per Vehicle After'] = '-'
        result['Workcenter No. After'] = '-'
        result['Workcenter Name After'] = '-'
        result['Supplier Name After'] = '-'
        result['Localization After'] = '-'

    # Поля, которые будут заполнены позже
    result['Batch fact'] = ''
    result['Change Date'] = ''
    result['Quantity batches in SS'] = ''
    result['Configuration for old parts using out'] = ''
    result['Batches for old parts using out'] = ''
    result['Transmission'] = ''
    result['Quantity per Box Before'] = ''
    result['Quantity per Box After'] = ''
    result['Box Before (L-W-H) mm'] = ''
    result['Box After (L-W-H) mm'] = ''
    result['Pallet Before (L-W-H) mm'] = ''
    result['Pallet After (L-W-H) mm'] = ''
    result['Comments'] = ''

    return result


def find_pairs(df_bp: pd.DataFrame) -> List[Dict]:
    """
    Шаг 2: Поиск пар Before/After деталей
    
    Правила:
    - Delete → всегда Before
    - Add без пары → After
    - Add + Delete → Add = After, Delete = Before
    - Add + Replace → Add = Before, Replace = After
    - Add + Update → Add = Before, Update = After
    """
    df_bp = df_bp.copy()
    df_bp['_direction'] = df_bp.apply(classify_row_before_after, axis=1)

    before_rows = df_bp[df_bp['_direction'] == 'Before'].copy()
    after_rows = pd.DataFrame()
    unknown_rows = df_bp[df_bp['_direction'] == 'Unknown'].copy()

    add_as_before = []
    temp_after = []

    for _, unknown in unknown_rows.iterrows():
        update_type = safe_str_convert(unknown.get('Update Type', ''))
        part_name = safe_str_convert(unknown.get('Part Name (RUS)', ''))

        if update_type == 'Add':
            match_delete = before_rows[
                before_rows['Part Name (RUS)'].astype(str).str.strip() == part_name
            ]
            if not match_delete.empty:
                temp_after.append(unknown)
                continue

            match_replace_update = unknown_rows[
                (unknown_rows['Update Type'].isin(['Replace', 'Update'])) &
                (unknown_rows['Part Name (RUS)'].astype(str).str.strip() == part_name)
            ]
            if not match_replace_update.empty:
                add_as_before.append(unknown)
            else:
                temp_after.append(unknown)

        elif update_type in ['Replace', 'Update']:
            match_add = unknown_rows[
                (unknown_rows['Update Type'] == 'Add') &
                (unknown_rows['Part Name (RUS)'].astype(str).str.strip() == part_name)
            ]
            if not match_add.empty:
                temp_after.append(unknown)
            else:
                temp_after.append(unknown)

    if add_as_before:
        add_before_df = pd.DataFrame(add_as_before)
        if '_direction' in add_before_df.columns:
            add_before_df = add_before_df.drop(columns=['_direction'])
        before_rows = pd.concat([before_rows, add_before_df], ignore_index=True)

    if temp_after:
        after_rows = pd.DataFrame(temp_after)
        if '_direction' in after_rows.columns:
            after_rows = after_rows.drop(columns=['_direction'])

    # Создаем пары
    pairs = []
    used_before = set()
    used_after = set()

    # Приоритет 1: одинаковый Part No.
    for _, before_row in before_rows.iterrows():
        before_part_no = safe_str_convert(before_row.get('Part No.', ''))
        if before_part_no == '' or before_part_no == '-':
            continue
        if after_rows.empty:
            break
        matches = after_rows[
            (after_rows['Part No.'].astype(str).str.strip() == before_part_no) &
            (after_rows.index not in used_after)
        ]
        for _, after_row in matches.iterrows():
            pairs.append(create_pair_dict(before_row, after_row))
            used_before.add(before_row.name)
            used_after.add(after_row.name)
            break

    # Приоритет 2: одинаковый Part Name (RUS)
    for _, before_row in before_rows.iterrows():
        if before_row.name in used_before:
            continue
        before_part_name = safe_str_convert(before_row.get('Part Name (RUS)', ''))
        if before_part_name == '' or before_part_name == '-':
            continue
        if after_rows.empty:
            break
        matches = after_rows[
            (after_rows['Part Name (RUS)'].astype(str).str.strip() == before_part_name) &
            (after_rows.index not in used_after)
        ]
        for _, after_row in matches.iterrows():
            pairs.append(create_pair_dict(before_row, after_row))
            used_before.add(before_row.name)
            used_after.add(after_row.name)
            break

    # Before без пары
    for _, before_row in before_rows.iterrows():
        if before_row.name not in used_before:
            pairs.append(create_pair_dict(before_row, None))
            used_before.add(before_row.name)

    # After без пары
    if not after_rows.empty:
        for _, after_row in after_rows.iterrows():
            if after_row.name not in used_after:
                pairs.append(create_pair_dict(None, after_row))
                used_after.add(after_row.name)

    return pairs


def user_input_for_single_bp(df_current: pd.DataFrame, bp_number: str) -> pd.DataFrame:
    """
    Интерактивный ввод Batch fact и Change Date для одного BP файла
    """
    print(f"\n--- Ввод данных для BP {bp_number} ---")
    print(f"Всего строк для обработки: {len(df_current)}")

    df_result = df_current.copy()
    total_rows = len(df_result)

    for counter, (idx, row) in enumerate(df_result.iterrows(), start=1):
        bom_product = safe_str_convert(row.get('BOM Product', ''))
        part_no_before = safe_str_convert(row.get('Part No. Before', ''))
        part_no_after = safe_str_convert(row.get('Part No. After', ''))

        print(f"\n[{counter}/{total_rows}] BOM: {bom_product}")
        print(f"  Before: {part_no_before}")
        print(f"  After: {part_no_after}")

        # Batch fact
        current_batch_fact = safe_str_convert(row.get('Batch fact', ''))
        user_input = input(f"  Batch fact [{current_batch_fact if current_batch_fact else 'пусто'}]: ").strip()
        if user_input:
            df_result.at[idx, 'Batch fact'] = user_input

        # Change Date
        current_change_date = safe_str_convert(row.get('Change Date', ''))
        user_input = input(f"  Change Date (ГГГГ-ММ-ДД) [{current_change_date if current_change_date else 'пусто'}]: ").strip()
        if user_input:
            df_result.at[idx, 'Change Date'] = user_input

    return df_result


def config_lookup_for_single_bp(df_current: pd.DataFrame, df_config: Optional[pd.DataFrame]) -> pd.DataFrame:
    """
    Поиск значений в конфигурационном файле для одного BP файла
    """
    if df_config is None or df_config.empty:
        print("  Конфигурационный файл не загружен. Пропускаем поиск.")
        return df_current

    df_result = df_current.copy()

    # Проверка наличия необходимых колонок
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

        # Quantity batches in SS
        quantity_vehicle_in_batch = config_matches.iloc[0].get('Quantity vehicle in batch')
        quantity_in_ss = row.get('Quantity in SS', 0)

        qty_in_ss = safe_float_convert(quantity_in_ss, 0.0)
        qty_vehicle = safe_float_convert(quantity_vehicle_in_batch, 1.0)

        if qty_vehicle > 0:
            qty_batches = round(qty_in_ss / qty_vehicle, 2)
            df_result.at[idx, 'Quantity batches in SS'] = qty_batches
            print(f"  {bom_product}: Quantity batches in SS = {qty_batches}")

        # Поиск по Batch fact
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

                    if 'Batch code' in df_config.columns:
                        batch_code_value = safe_str_convert(config_match.iloc[0].get('Batch code', ''))
                        if batch_code_value and batch_code_value != 'nan':
                            df_result.at[idx, 'Batches for old parts using out'] = batch_code_value

                    if 'Transmission' in df_config.columns:
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
    Обработка одного BP файла:
    1. Поиск пар Before/After
    2. Пользовательский ввод
    3. Поиск в конфигурации
    4. Загрузка файлов партий
    """
    print(f"\n{'=' * 60}")
    print(f"ОБРАБОТКА BP: {bp_number}")
    print(f"{'=' * 60}")

    # Шаг 1: Поиск пар
    print("\n[1/4] Поиск пар Before/After...")
    pairs = find_pairs(df_bp)
    print(f"  Найдено пар/строк: {len(pairs)}")

    if not pairs:
        print(f"  Предупреждение: Для {bp_number} не найдено пар!")
        return pd.DataFrame()

    df_current = pd.DataFrame(pairs)

    # Шаг 2: Пользовательский ввод
    if interactive:
        print("\n[2/4] Ввод данных пользователем...")
        df_current = user_input_for_single_bp(df_current, bp_number)

    # Шаг 3: Поиск в конфигурации
    print("\n[3/4] Поиск в конфигурационном файле...")
    df_current = config_lookup_for_single_bp(df_current, df_config)

    # Шаг 4: Загрузка файлов партий
    print("\n[4/4] Проверка файлов партий...")
    df_current = batch_file_loader_for_single_bp(df_current)

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
        'Localization After', 'Change Description', 'Change Solution', 'Comments',
        'Color Code', 'Color Name (RUS)'
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

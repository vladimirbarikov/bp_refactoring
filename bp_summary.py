#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль summary_breakpoint_table
Преобразует обработанные BP DataFrame в итоговый df_summary_breakpoint
"""

import os
import sys
import io
from typing import Dict, List, Optional

import pandas as pd

# Путь к сетевым папкам с файлами партий (при необходимости можно переопределить)
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
    
    Args:
        value: значение для преобразования
        default: значение по умолчанию при ошибке
    
    Returns:
        float - преобразованное значение или default
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
    
    Args:
        value: значение для преобразования
        default: значение по умолчанию при ошибке
    
    Returns:
        str - преобразованная строка или default
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
        print("Некоторые поля (Quantity batches in SS, Configuration for old parts using out, Transmission) будут пустыми")
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
        print("Закройте файл, если он открыт в Excel, и попробуйте снова.")
        return None
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
        print(f"Ошибка: Файл '{config_filename}' повреждён или имеет неверный формат: {e}")
        return None
    except ValueError as e:
        if "Excel file format cannot be determined" in str(e):
            print(f"Ошибка: Не удалось определить формат файла '{config_filename}'")
            print("Убедитесь, что файл имеет расширение .xlsx или .xls")
        else:
            print(f"Ошибка при загрузке файла конфигурации: {e}")
        return None
    except KeyError as e:
        print(f"Ошибка: В файле конфигурации отсутствует необходимая колонка: {e}")
        print("Проверьте, что файл содержит колонки: 'BOM Product', 'Quantity vehicle in batch', 'Batch code', 'Configuration', 'Transmission'")
        return None
    except Exception as e:
        print(f"Непредвиденная ошибка при загрузке файла конфигурации: {e}")
        print(f"Тип ошибки: {type(e).__name__}")
        return None


def classify_row_before_after(row: pd.Series) -> str:
    """
    Шаг 1: Классификация детали на 'Before' или 'After'
    Приоритет 1: колонка 'Change'
    Приоритет 2: колонка 'Update Type'
    
    Правила:
    - Delete → всегда Before
    - Add, Replace, Update → требуют анализа пар (Unknown)
    """
    # Приоритет 1: явное указание в колонке Change
    change_val = safe_str_convert(row.get('Change', ''))
    if change_val == 'Before Change':
        return 'Before'
    elif change_val == 'After Change':
        return 'After'

    # Приоритет 2: анализ Update Type
    update_type = safe_str_convert(row.get('Update Type', ''))

    # Delete всегда Before
    if update_type == 'Delete':
        return 'Before'
    # Add, Replace, Update требуют поиска пар
    elif update_type in ['Add', 'Replace', 'Update']:
        return 'Unknown'
    else:
        return 'Unknown'


def create_pair_dict(before_row: Optional[pd.Series], after_row: Optional[pd.Series]) -> Dict:
    """Создает словарь для одной строки итоговой таблицы из пары Before/After"""
    result = {}

    # Базовые поля (общие для BP)
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
    
    Возвращает список словарей, где каждый словарь - одна строка итоговой таблицы
    """
    # Сначала классифицируем все строки
    df_bp = df_bp.copy()
    df_bp['_direction'] = df_bp.apply(classify_row_before_after, axis=1)

    # Delete строки сразу идут в Before
    before_rows = df_bp[df_bp['_direction'] == 'Before'].copy()
    # После обработки Unknown, After строки будут собираться сюда
    after_rows = pd.DataFrame()
    
    # Все строки, требующие анализа (Add, Replace, Update)
    unknown_rows = df_bp[df_bp['_direction'] == 'Unknown'].copy()
    
    # Временные хранилища для Add, которые станут Before
    add_as_before = []
    # Временные хранилища для остальных
    temp_after = []

    # Проходим по всем Unknown строкам для определения их роли
    for _, unknown in unknown_rows.iterrows():
        update_type = safe_str_convert(unknown.get('Update Type', ''))
        part_name = safe_str_convert(unknown.get('Part Name (RUS)', ''))
        
        if update_type == 'Add':
            # Проверяем, есть ли Delete с таким же Part Name
            match_delete = before_rows[
                before_rows['Part Name (RUS)'].astype(str).str.strip() == part_name
            ]
            if not match_delete.empty:
                # Add + Delete → Add = After, Delete уже в before_rows
                temp_after.append(unknown)
                continue
            
            # Проверяем, есть ли Replace или Update с таким же Part Name
            match_replace_update = unknown_rows[
                (unknown_rows['Update Type'].isin(['Replace', 'Update'])) &
                (unknown_rows['Part Name (RUS)'].astype(str).str.strip() == part_name)
            ]
            if not match_replace_update.empty:
                # Add + Replace/Update → Add = Before
                add_as_before.append(unknown)
            else:
                # Add без пары → After
                temp_after.append(unknown)
        
        elif update_type in ['Replace', 'Update']:
            # Проверяем, есть ли Add с таким же Part Name
            match_add = unknown_rows[
                (unknown_rows['Update Type'] == 'Add') &
                (unknown_rows['Part Name (RUS)'].astype(str).str.strip() == part_name)
            ]
            if not match_add.empty:
                # Replace/Update с парой Add → After
                temp_after.append(unknown)
            else:
                # Replace/Update без пары (по правилам не должно быть, но на всякий случай - After)
                temp_after.append(unknown)

    # Добавляем Add как Before (для пар с Replace/Update)
    if add_as_before:
        add_before_df = pd.DataFrame(add_as_before)
        # Удаляем колонку _direction, если она есть
        if '_direction' in add_before_df.columns:
            add_before_df = add_before_df.drop(columns=['_direction'])
        before_rows = pd.concat([before_rows, add_before_df], ignore_index=True)

    # Формируем after_rows
    if temp_after:
        after_rows = pd.DataFrame(temp_after)
        if '_direction' in after_rows.columns:
            after_rows = after_rows.drop(columns=['_direction'])

    # Создаем пары Before-After
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

    # Обрабатываем Before без пары
    for _, before_row in before_rows.iterrows():
        if before_row.name not in used_before:
            pairs.append(create_pair_dict(before_row, None))
            used_before.add(before_row.name)

    # Обрабатываем After без пары
    if not after_rows.empty:
        for _, after_row in after_rows.iterrows():
            if after_row.name not in used_after:
                pairs.append(create_pair_dict(None, after_row))
                used_after.add(after_row.name)

    return pairs


def user_input_batch_fact_and_change_date(df_summary: pd.DataFrame) -> pd.DataFrame:
    """
    Интерактивный ввод Batch fact и Change Date для каждой строки
    """
    print("\n" + "=" * 60)
    print("ВВОД ДАННЫХ: Batch fact и Change Date")
    print("=" * 60)
    print("Для каждой строки BP можно указать:")
    print("  - Batch fact (фактическая партия)")
    print("  - Change Date (дата изменения)")
    print("Если данные неизвестны - оставьте поле пустым и нажмите Enter")

    df_result = df_summary.copy()

    for idx, row in df_result.iterrows():
        bp_no = safe_str_convert(row.get('BP_No', ''))
        bom_product = safe_str_convert(row.get('BOM Product', ''))
        part_no_after = safe_str_convert(row.get('Part No. After', ''))

        print(f"\n--- BP: {bp_no} | BOM Product: {bom_product} | Part No. After: {part_no_after} ---")

        # Batch fact
        current_batch_fact = safe_str_convert(row.get('Batch fact', ''))
        print(f"Текущее Batch fact: {current_batch_fact if current_batch_fact else '(пусто)'}")
        user_input = input("Введите Batch fact (или Enter чтобы оставить пустым): ").strip()
        if user_input:
            df_result.at[idx, 'Batch fact'] = user_input

        # Change Date
        current_change_date = safe_str_convert(row.get('Change Date', ''))
        print(f"Текущее Change Date: {current_change_date if current_change_date else '(пусто)'}")
        user_input = input("Введите Change Date (формат ГГГГ-ММ-ДД или Enter): ").strip()
        if user_input:
            df_result.at[idx, 'Change Date'] = user_input

    return df_result


def config_lookup(df_summary: pd.DataFrame, df_config: Optional[pd.DataFrame]) -> pd.DataFrame:
    """
    Поиск значений в конфигурационном файле:
    - Quantity batches in SS
    - Configuration for old parts using out
    - Batches for old parts using out
    - Transmission
    """
    if df_config is None or df_config.empty:
        print("\nКонфигурационный файл не загружен. Пропускаем поиск значений.")
        return df_summary

    print("\n" + "=" * 60)
    print("ПОИСК ЗНАЧЕНИЙ В КОНФИГУРАЦИОННОМ ФАЙЛЕ")
    print("=" * 60)

    df_result = df_summary.copy()

    # Проверка наличия необходимых колонок в df_config
    required_config_cols = ['BOM Product', 'Quantity vehicle in batch']
    optional_config_cols = ['Batch code', 'Configuration', 'Transmission']

    missing_cols = [col for col in required_config_cols if col not in df_config.columns]
    if missing_cols:
        print(f"Ошибка: В конфигурационном файле отсутствуют обязательные колонки: {missing_cols}")
        print("Поиск значений отменён.")
        return df_result

    # Приводим колонки df_config к строковому типу (только если они существуют)
    df_config['BOM Product'] = df_config['BOM Product'].astype(str).str.strip()

    for col in optional_config_cols:
        if col in df_config.columns:
            df_config[col] = df_config[col].astype(str).str.strip()

    for idx, row in df_result.iterrows():
        bom_product = safe_str_convert(row.get('BOM Product', ''))
        batch_fact = safe_str_convert(row.get('Batch fact', ''))

        if bom_product == '' or bom_product == '-':
            continue

        # Поиск всех строк с таким BOM Product
        config_matches = df_config[df_config['BOM Product'] == bom_product]

        if config_matches.empty:
            continue

        # 1. Quantity vehicle in batch (берем первое значение)
        quantity_vehicle_in_batch = config_matches.iloc[0].get('Quantity vehicle in batch')
        quantity_in_ss = row.get('Quantity in SS', 0)

        # Используем safe_float_convert
        qty_in_ss = safe_float_convert(quantity_in_ss, 0.0)
        qty_vehicle = safe_float_convert(quantity_vehicle_in_batch, 1.0)

        if qty_vehicle > 0:
            qty_batches = round(qty_in_ss / qty_vehicle, 2)
            df_result.at[idx, 'Quantity batches in SS'] = qty_batches
            print(f"  {bom_product}: Quantity batches in SS = {qty_batches} (={qty_in_ss}/{qty_vehicle})")
        elif qty_vehicle == 0:
            print(f"  Предупреждение: Quantity vehicle in batch для {bom_product} = 0 (деление на ноль)")

        # 2. Поиск по Batch fact (только если есть колонки Batch code, Configuration, Transmission)
        if batch_fact and batch_fact != '' and batch_fact != '-':
            if 'Batch code' in df_config.columns and 'Configuration' in df_config.columns:
                # Извлекаем первые 3 символа Batch fact для поиска
                batch_prefix = batch_fact[:3] if len(batch_fact) >= 3 else batch_fact

                # Ищем по BOM Product и Batch code (первые 3 символа)
                config_match = config_matches[
                    config_matches['Batch code'].str.startswith(batch_prefix, na=False)
                ]

                if not config_match.empty:
                    # Configuration for old parts using out
                    config_value = safe_str_convert(config_match.iloc[0].get('Configuration', ''))
                    if config_value and config_value != '' and config_value != 'nan':
                        df_result.at[idx, 'Configuration for old parts using out'] = config_value

                    # Batches for old parts using out
                    if 'Batch code' in df_config.columns:
                        batch_code_value = safe_str_convert(config_match.iloc[0].get('Batch code', ''))
                        if batch_code_value and batch_code_value != '' and batch_code_value != 'nan':
                            df_result.at[idx, 'Batches for old parts using out'] = batch_code_value

                    # Transmission
                    if 'Transmission' in df_config.columns:
                        transmission_value = safe_str_convert(config_match.iloc[0].get('Transmission', ''))
                        if transmission_value and transmission_value != '' and transmission_value != 'nan':
                            df_result.at[idx, 'Transmission'] = transmission_value

                    print(f"  {bom_product} | Batch fact: {batch_fact} → найдена конфигурация")
                else:
                    print(f"  {bom_product} | Batch fact: {batch_fact} → конфигурация не найдена")
            else:
                print("  Предупреждение: В конфигурационном файле отсутствуют колонки 'Batch code' или 'Configuration'")

    return df_result


def load_batch_file(batch_name: str, search_path: str = BATCH_FILES_PATH) -> Optional[pd.DataFrame]:
    """
    Загрузка Excel файла партии
    
    Args:
        batch_name: название партии (например, 'RKV2029')
        search_path: путь к папке с файлами партий
    
    Returns:
        DataFrame с данными партии или None
    """
    batch_name = safe_str_convert(batch_name)

    if not batch_name or batch_name == '' or batch_name == '-':
        print("  Пропуск: название партии не указано")
        return None

    # Генерация имени файла: '2029 RKV.xlsx' из 'RKV2029'
    if len(batch_name) >= 4:
        file_number = batch_name[3:]  # '2029' из 'RKV2029'
        file_prefix = batch_name[:3]   # 'RKV' из 'RKV2029'
        filename = f"{file_number} {file_prefix}.xlsx"
    else:
        filename = f"{batch_name}.xlsx"

    full_path = os.path.join(search_path, filename)

    if not os.path.exists(full_path):
        print(f"  Файл партии не найден: {full_path}")
        return None

    try:
        df_batch = pd.read_excel(full_path)
        print(f"  Файл партии загружен: {filename} ({df_batch.shape[0]} строк)")
        return df_batch
    except FileNotFoundError:
        print(f"  Ошибка: Файл '{full_path}' не найден")
        return None
    except PermissionError:
        print(f"  Ошибка: Нет прав для чтения файла '{full_path}'")
        print("  Закройте файл, если он открыт в Excel, и попробуйте снова.")
        return None
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
        print(f"  Ошибка: Файл '{filename}' повреждён: {e}")
        return None
    except Exception as e:
        print(f"  Непредвиденная ошибка при загрузке файла '{filename}': {e}")
        return None


def batch_file_loader_placeholder(df_summary: pd.DataFrame) -> pd.DataFrame:
    """
    Загрузка Excel файлов партий для заполнения информации о коробках и паллетах
    
    Логика работы (будет реализована позже):
    1. Для каждой строки в df_summary проверяем наличие Batch fact
    2. Если Batch fact заполнен, вызываем load_batch_file() для загрузки файла партии
    3. Из загруженного файла извлекаем:
       - Quantity per Box Before/After
       - Box Before/After (L-W-H) mm
       - Pallet Before/After (L-W-H) mm
    4. Заполняем соответствующие колонки в df_summary
    
    Формат файлов партий:
    - Имя файла: '{номер} {префикс}.xlsx' (например, '2029 RKV.xlsx' для партии 'RKV2029')
    - Путь: \\hmmr_share\LD\Custom Clearance\Поставки\Серийный KD parts
    
    TODO: Реализовать парсинг структуры Excel файла партии:
    - Определить, на каких листах и в каких колонках находятся данные
    - Связать детали Before/After с данными из файла
    - Заполнить целевые колонки
    
    P.S. Этот функционал будет реализован после получения образцов файлов партий
    """
    print("\n" + "=" * 60)
    print("ЗАГРУЗКА ФАЙЛОВ ПАРТИЙ")
    print("=" * 60)
    print("ВНИМАНИЕ: Этот функционал находится в разработке.")
    print("На данном этапе выполняется только проверка наличия файлов партий.")
    print("Поля 'Quantity per Box', 'Box (L-W-H) mm', 'Pallet (L-W-H) mm' будут пустыми.")
    print("Вы сможете заполнить их позже в Excel файле.\n")

    df_result = df_summary.copy()

    # Проверяем наличие Batch fact и пытаемся загрузить файлы
    for _, row in df_result.iterrows():
        batch_fact = safe_str_convert(row.get('Batch fact', ''))
        bp_no = safe_str_convert(row.get('BP_No', ''))

        if batch_fact and batch_fact != '' and batch_fact != '-':
            print(f"Проверка файла для партии: {batch_fact}")
            df_batch = load_batch_file(batch_fact)

            if df_batch is not None:
                print(f"  → Файл партии '{batch_fact}' успешно загружен")
                # TODO: Здесь будет логика извлечения данных из df_batch
                # и заполнения колонок:
                # - Quantity per Box Before/After
                # - Box Before/After (L-W-H) mm
                # - Pallet Before/After (L-W-H) mm
            else:
                print(f"  → Файл для партии '{batch_fact}' не найден или не может быть загружен")
        else:
            print(f"Пропуск: для BP {bp_no} не указан Batch fact")

    print("\n" + "=" * 60)
    print("Завершена проверка файлов партий.")
    print("Для заполнения данных о коробках и паллетах используйте Excel файл.")
    print("=" * 60)

    return df_result


def create_summary_breakpoint_table(
    processed_results: Dict[str, pd.DataFrame]
) -> pd.DataFrame:
    """
    Основная функция: создает итоговый df_summary_breakpoint из processed_results
    
    Args:
        processed_results: словарь вида {bp_number: dataframe} из bp_refactoring
    
    Returns:
        pd.DataFrame - итоговая таблица summary_breakpoint
    """
    all_pairs = []

    print("\n" + "=" * 60)
    print("ОБРАБОТКА BP ФАЙЛОВ ДЛЯ SUMMARY BREAKPOINT")
    print("=" * 60)

    for bp_number, df_bp in processed_results.items():
        print(f"\nОбработка: {bp_number} ({len(df_bp)} строк)")

        # Находим пары Before/After
        pairs = find_pairs(df_bp)
        print(f"  Найдено пар/строк: {len(pairs)}")

        all_pairs.extend(pairs)

    # Создаем DataFrame из пар
    df_summary = pd.DataFrame(all_pairs)

    # Упорядочиваем колонки согласно спецификации
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

    # Добавляем отсутствующие колонки
    for col in column_order:
        if col not in df_summary.columns:
            df_summary[col] = ''

    df_summary = df_summary[column_order]

    return df_summary


def main(processed_results: Dict[str, pd.DataFrame], interactive: bool = True) -> pd.DataFrame:
    """
    Главная функция модуля summary_breakpoint_table
    
    Args:
        processed_results: словарь из bp_refactoring (bp_number -> dataframe)
        interactive: если True, запрашивает пользовательский ввод
    
    Returns:
        pd.DataFrame - итоговый df_summary_breakpoint
    """
    clear_screen()

    print("""
        ╔══════════════════════════════════════════════════════════════╗
        ║           SUMMARY BREAKPOINT TABLE GENERATOR v1.0            ║
        ║        Формирование итоговой таблицы из BP файлов            ║
        ╚══════════════════════════════════════════════════════════════╝
        """)

    print(f"Найдено BP файлов для обработки: {len(processed_results)}")

    # Шаг 1: Загрузка конфигурационного файла
    print("\n" + "=" * 60)
    print("ШАГ 1: Загрузка конфигурационного файла")
    print("=" * 60)
    df_config = load_configuration_file('configuration.xlsx')
    wait_for_user()

    # Шаг 2: Создание базовой таблицы (поиск пар Before/After)
    print("\n" + "=" * 60)
    print("ШАГ 2: Создание базовой таблицы summary_breakpoint")
    print("=" * 60)
    df_summary = create_summary_breakpoint_table(processed_results)
    print(f"\nСоздана таблица: {df_summary.shape[0]} строк × {df_summary.shape[1]} колонок")
    wait_for_user()

    # Шаг 3: Пользовательский ввод Batch fact и Change Date
    if interactive:
        print("\n" + "=" * 60)
        print("ШАГ 3: Ввод данных пользователем")
        print("=" * 60)
        df_summary = user_input_batch_fact_and_change_date(df_summary)
        wait_for_user()

    # Шаг 4: Поиск в конфигурационном файле
    print("\n" + "=" * 60)
    print("ШАГ 4: Поиск значений в конфигурационном файле")
    print("=" * 60)
    df_summary = config_lookup(df_summary, df_config)
    wait_for_user()

    # Шаг 5: Загрузка файлов партий (заглушка)
    print("\n" + "=" * 60)
    print("ШАГ 5: Загрузка файлов партий")
    print("=" * 60)
    df_summary = batch_file_loader_placeholder(df_summary)
    wait_for_user()

    # Итоги
    clear_screen()
    print("\n" + "=" * 60)
    print("ИТОГОВАЯ ТАБЛИЦА SUMMARY BREAKPOINT")
    print("=" * 60)
    print(f"Сформировано строк: {len(df_summary)}")
    print(f"Колонок: {len(df_summary.columns)}")

    return df_summary


# Пример использования при запуске модуля напрямую
if __name__ == "__main__":
    # Этот блок будет использоваться для тестирования
    # В реальной жизни processed_results приходит из bp_refactoring

    print("Этот модуль предназначен для импорта из bp_refactoring")
    print("Использование:")
    print("  from summary_breakpoint_table import main")
    print("  df_summary = main(processed_results)")

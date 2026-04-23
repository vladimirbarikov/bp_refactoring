#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BP Refactoring Tool - Пошаговая обработка Excel файлов
Использование: python bp_refactoring.py
"""

import os
import re
import sys
import warnings
from datetime import datetime

import pandas as pd
from pandas.errors import EmptyDataError, ParserError
from openpyxl.utils.exceptions import InvalidFileException

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def clear_screen():
    """Очистка экрана консоли"""
    os.system('cls' if os.name == 'nt' else 'clear')


def wait_for_user(prompt="\nНажмите Enter для продолжения..."):
    """Ожидание нажатия Enter"""
    input(prompt)


def print_step_header(step_num, total_steps, description):
    """Вывод заголовка шага"""
    print("\n" + "=" * 60)
    print(f"ШАГ {step_num}/{total_steps}: {description}")
    print("=" * 60)


def load_excel_file(filename, description="файла"):
    """Загрузка Excel файла с проверкой существования"""
    if not os.path.exists(filename):
        print(f"Ошибка: {description} '{filename}' не найден в текущей папке!")
        print(f"Текущая директория: {os.getcwd()}")
        return None

    try:
        df = pd.read_excel(filename)
        print(f"{description} '{filename}' загружен успешно!")
        print(f"Размер: {df.shape[0]} строк × {df.shape[1]} колонок")
        return df
    except FileNotFoundError:
        print(f"Ошибка: {description} '{filename}' не найден (ошибка FileNotFoundError)")
        return None
    except PermissionError:
        print(f"Ошибка: Нет прав для чтения {description.lower()} '{filename}'")
        print("Закройте файл, если он открыт в Excel, и попробуйте снова.")
        return None
    except (EmptyDataError, ParserError) as e:
        print(f"Ошибка: {description} '{filename}' повреждён или имеет неверный формат: {e}")
        return None
    except InvalidFileException:
        print(f"Ошибка: {description} '{filename}' не является корректным Excel файлом")
        return None
    except ValueError as e:
        if "Excel file format cannot be determined" in str(e):
            print(f"Ошибка: Не удалось определить формат {description.lower()} '{filename}'")
            print("Убедитесь, что файл имеет расширение .xlsx или .xls")
        else:
            print(f"Ошибка при загрузке {description.lower()} '{filename}': {e}")
        return None
    except Exception as e:
        print(f"Непредвиденная ошибка при загрузке {description.lower()} '{filename}': {e}")
        print(f"Тип ошибки: {type(e).__name__}")
        return None


def save_excel_file(df, filename):
    """Сохранение Excel файла с проверкой на дубликаты"""
    if df is None or df.empty:
        print(f"Внимание: Нет данных для сохранения в {filename}")
        return False

    if os.path.exists(filename):
        print(f"Файл {filename} уже существует!")
        name, ext = filename.rsplit('.', 1)
        new_filename = f"{name}_v2.{ext}"
        print(f"Сохраняем как {new_filename}")
        filename = new_filename

    try:
        df.to_excel(filename, index=False)
        print(f"Файл сохранен: {filename}")
        return True
    except PermissionError:
        print(f"Ошибка: Нет прав для записи в файл '{filename}'")
        print("Закройте файл, если он открыт в Excel, и попробуйте снова.")
        return False
    except OSError as e:
        print(f"Ошибка при сохранении: {e}")
        return False
    except Exception as e:
        print(f"Непредвиденная ошибка при сохранении: {e}")
        return False


def show_dataframe_preview(df, step_name, max_rows=5, focus_columns=None, max_colwidth=40):
    """
    Отображает первые строки DataFrame для визуального контроля
    step_name: название шага
    max_rows: количество строк для отображения
    focus_columns: список колонок для отображения (если None - показывает первые 5)
    max_colwidth: максимальная ширина содержимого колонки в символах
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

    # Обрезаем длинные текстовые значения
    for col in preview_df.columns:
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


def filter_chinese_lines(text):
    """Фильтрация китайских иероглифов"""
    if not isinstance(text, str):
        return text

    chinese_pattern = re.compile(r'[\u4e00-\u9fff]')

    if '\n' in text:
        lines = text.split('\n')
        chinese_lines = [line.strip() for line in lines if chinese_pattern.search(line)]
        return ', '.join(chinese_lines) if chinese_lines else text
    else:
        last_chinese_pos = -1
        for i, char in enumerate(text):
            if chinese_pattern.match(char):
                last_chinese_pos = i

        if last_chinese_pos != -1:
            return text[:last_chinese_pos + 1]
        else:
            return text


def extract_parentheses_content(text):
    """Извлечение содержимого скобок"""
    if pd.isna(text):
        return text

    text_str = str(text)
    matches = re.findall(r'[（(](.*?)[）)]', text_str)

    if matches:
        return ' '.join(matches)

    return text_str


def fill_empty_values_with_dash(df, columns):
    """
    Заполняет пустые значения в указанных колонках на '-'
    Возвращает список колонок, в которых были замены
    """
    columns_with_replacements = []

    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str)

            empty_mask = df[col].isna() | (df[col].str.strip() == '') | (df[col].str.strip() == 'nan') | (df[col].str.strip() == 'None')
            empty_count = empty_mask.sum()

            if empty_count > 0:
                print(f"  Колонка '{col}': ячейки без данных - {empty_count}. Заполняем '-'.")
                df.loc[empty_mask, col] = '-'
                columns_with_replacements.append(col)
            else:
                print(f"  Колонка '{col}': все данные заполнены. Оставляем без изменений.")
        else:
            print(f"  Колонка '{col}': отсутствует в данных.")

    return columns_with_replacements


def get_unique_non_empty_values(series, column_name):
    """Получает уникальные значения из серии, исключая '-' и пустые значения"""
    series = series.astype(str)
    unique_values = series.dropna().unique()
    filtered_values = [v for v in unique_values if v != '-' and v != 'nan' and str(v).strip() != '']

    if len(filtered_values) == 0:
        print(f"  Колонка '{column_name}': все значения равны '-' или пустые. Перевод не требуется.")

    return filtered_values


def interactive_translation(data, field_name, examples=None):
    """
    Интерактивный ввод переводов
    data: список уникальных значений для перевода
    field_name: название поля (для вывода)
    examples: примеры переводов (словарь)
    """
    if not data or len(data) == 0:
        return {}

    translations = {}

    print(f"\nПеревод {field_name}:")
    print(f"Найдено {len(data)} уникальных значений")

    if examples:
        print("Примеры переводов (можно использовать как шаблон):")
        for ch, ru in examples.items():
            print(f"     {ch[:50]}... → {ru[:50]}...")

    print("\nИнструкция:")
    print("  • Введите перевод и нажмите Enter → оригинальный текст будет заменён на перевод")
    print("  • Нажмите Enter без перевода → текст останется оригинальным (без изменений)")

    for i, value in enumerate(data, 1):
        if pd.isna(value) or value == '':
            translations[value] = value  # оставляем как есть
            continue

        print(f"\n[{i}/{len(data)}] Оригинал: {value}")
        try:
            user_input = input("Введите перевод (или просто Enter чтобы оставить оригинал): ").strip()
        except EOFError:
            print("\nОбнаружен конец ввода. Оставляем оригинальные значения.")
            translations[value] = value
            for remaining in data[i:]:
                translations[remaining] = remaining
            break
        except KeyboardInterrupt:
            print("\nВвод прерван. Оставляем оригинальные значения.")
            translations[value] = value
            for remaining in data[i:]:
                translations[remaining] = remaining
            break

        if user_input == '':
            # Пустой ввод - оставляем оригинал
            translations[value] = value
            print(f"  → Оставляем оригинал: {value}")
        else:
            # Есть перевод - заменяем
            translations[value] = user_input
            print(f"  → Заменяем на: {user_input}")

    return translations


def find_bp_files():
    """Поиск BP файлов в текущей папке"""
    bp_files = []
    try:
        for file in os.listdir('.'):
            if re.match(r'^BP.*\.xlsx$', file) and not file.startswith('~$'):
                bp_files.append(file)
    except PermissionError:
        print("Ошибка: Нет прав для чтения текущей директории")
        return []
    except OSError as e:
        print(f"Ошибка при доступе к директории: {e}")
        return []

    bp_files.sort()
    return bp_files


def confirm_step(step_name):
    """
    Запрашивает у пользователя подтверждение после выполнения шага
    Возвращает True если нужно продолжить, False если нужно повторить шаг
    """
    print(f"\n  Шаг '{step_name}' выполнен.")
    user_input = input(
        "\nПроверьте результат. Если всё корректно, нажмите Enter. Если нужно повторить шаг, введите 'retry': "
    ).strip().lower()

    if user_input == 'retry':
        print(f"Повторяем шаг '{step_name}'...\n")
        return False
    else:
        print("Продолжаем...\n")
        return True


def process_bp_file(bp_filename, df_bom):
    """
    Обработка одного BP файла
    """
    print(f"\nОбработка файла: {bp_filename}")

    bp_number = bp_filename.replace('.xlsx', '')

    # Шаг 1: Загрузка BP файла
    print_step_header(1, 13, "Загрузка BP файла")
    df_bp = load_excel_file(bp_filename, "BP файла")
    if df_bp is None:
        return None
    show_dataframe_preview(df_bp, "Загрузка исходного BP файла")
    if not confirm_step("Загрузка BP файла"):
        return process_bp_file(bp_filename, df_bom)

    # Шаг 2: Выбор нужных колонок
    print_step_header(2, 13, "Выбор нужных колонок")
    bp_columns_to_keep = [
        'Change', 'BOM Product', 'Update Type', 'Part No.', 'Part Name(CHN)',
        'Quantity', 'Supplier Name', 'Change Description', 'Solution',
        'Color Code', 'Color Name', 'Production Part Disposal',
        'Interchangeable', 'In Stock', 'New Part Available Date',
        'Workcenter No.', 'Workcenter Name',
    ]

    available_cols = [col for col in bp_columns_to_keep if col in df_bp.columns]
    missing_cols = set(bp_columns_to_keep) - set(available_cols)
    if missing_cols:
        print(f"Внимание: Отсутствуют колонки: {missing_cols}")

    if not available_cols:
        print("Ошибка: В файле нет ни одной необходимой колонки")
        return None

    df_bp_new = df_bp[available_cols].copy()
    df_bp_new['BP_No'] = bp_number
    show_dataframe_preview(df_bp_new, "Выбор нужных колонок")
    if not confirm_step("Выбор нужных колонок"):
        return process_bp_file(bp_filename, df_bom)

    # Шаг 3: Заполнение пустых значений
    print_step_header(3, 13, "Заполнение пустых значений")
    # Заполняем пустые значения и запоминаем, в каких колонках были замены
    columns_with_replacements = fill_empty_values_with_dash(df_bp_new, available_cols)

    # Определяем колонки для preview: базовые + колонки с заменами
    base_columns = ['BOM Product', 'Part No.', 'Part Name(CHN)']
    # Объединяем, убираем дубликаты и сохраняем порядок
    preview_columns = []
    for col in base_columns:
        if col in df_bp_new.columns and col not in preview_columns:
            preview_columns.append(col)
    for col in columns_with_replacements:
        if col in df_bp_new.columns and col not in preview_columns:
            preview_columns.append(col)

    if preview_columns:
        show_dataframe_preview(df_bp_new, "Заполнение пустых значений",
                            focus_columns=preview_columns)
    else:
        show_dataframe_preview(df_bp_new, "Заполнение пустых значений",
                            focus_columns=['BOM Product', 'Part No.', 'Part Name(CHN)'])

    if not confirm_step("Заполнение пустых значений"):
        return process_bp_file(bp_filename, df_bom)

    # Шаг 4: Перевод названий деталей
    while True:
        print_step_header(4, 13, "Перевод названий деталей")
        if 'Part Name(CHN)' in df_bp_new.columns:
            unique_parts = get_unique_non_empty_values(df_bp_new['Part Name(CHN)'], 'Part Name(CHN)')

            if len(unique_parts) > 0:
                part_translations = interactive_translation(unique_parts, "названий деталей")
                df_bp_new['Part Name (RUS)'] = df_bp_new['Part Name(CHN)'].map(part_translations).fillna('-')
                df_bp_new = df_bp_new.drop(['Part Name(CHN)'], axis=1)
            else:
                print("  Колонка 'Part Name(CHN)': все значения равны '-' или пустые. Перевод не требуется.")
                df_bp_new['Part Name (RUS)'] = '-'
                df_bp_new = df_bp_new.drop(['Part Name(CHN)'], axis=1)
            show_dataframe_preview(df_bp_new, "Перевод названий деталей (после)",
                                  focus_columns=['Change', 'BOM Product', 'Part No.', 'Part Name (RUS)'])

        if confirm_step("Перевод названий деталей"):
            break

    # Шаг 5: Перевод поставщиков
    while True:
        print_step_header(5, 13, "Перевод названий поставщиков")
        if 'Supplier Name' in df_bp_new.columns:
            unique_suppliers = get_unique_non_empty_values(df_bp_new['Supplier Name'], 'Supplier Name')

            if len(unique_suppliers) > 0:
                supplier_translations = interactive_translation(unique_suppliers, "поставщиков")
                df_bp_new['Supplier Name (RUS)'] = df_bp_new['Supplier Name'].map(supplier_translations).fillna('-')
                df_bp_new = df_bp_new.drop(['Supplier Name'], axis=1)
            else:
                print("  Колонка 'Supplier Name': все значения равны '-' или пустые. Перевод не требуется.")
                df_bp_new['Supplier Name (RUS)'] = '-'
                df_bp_new = df_bp_new.drop(['Supplier Name'], axis=1)
            show_dataframe_preview(df_bp_new, "Перевод поставщиков (после)",
                                  focus_columns=['Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Supplier Name (RUS)'])

        if confirm_step("Перевод названий поставщиков"):
            break

    # Шаг 6: Фильтрация китайских символов
    print_step_header(6, 13, "Фильтрация китайских символов")
    if 'Change Description' in df_bp_new.columns:
        df_bp_new['Change Description'] = df_bp_new['Change Description'].apply(filter_chinese_lines)
        print("  Колонка 'Change Description': фильтрация выполнена")
    if 'Solution' in df_bp_new.columns:
        df_bp_new['Solution'] = df_bp_new['Solution'].apply(filter_chinese_lines)
        print("  Колонка 'Solution': фильтрация выполнена")
    show_dataframe_preview(df_bp_new, "Фильтрация китайских символов",
                          focus_columns=['Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Change Description', 'Solution'])
    if not confirm_step("Фильтрация китайских символов"):
        return process_bp_file(bp_filename, df_bom)

    # Шаг 7: Перевод описания
    while True:
        print_step_header(7, 13, "Перевод описания изменений")
        if 'Change Description' in df_bp_new.columns:
            unique_descs = get_unique_non_empty_values(df_bp_new['Change Description'], 'Change Description')

            if len(unique_descs) > 0:
                desc_translations = interactive_translation(unique_descs, "описаний изменений")
                df_bp_new['Change Description (RUS)'] = df_bp_new['Change Description'].map(desc_translations).fillna('-')
                df_bp_new = df_bp_new.drop(['Change Description'], axis=1)
            else:
                print("  Колонка 'Change Description': все значения равны '-' или пустые. Перевод не требуется.")
                df_bp_new['Change Description (RUS)'] = '-'
                df_bp_new = df_bp_new.drop(['Change Description'], axis=1)
            show_dataframe_preview(df_bp_new, "Перевод описания (после)",
                                  focus_columns=['Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Change Description (RUS)'])

        if confirm_step("Перевод описания изменений"):
            break

    # Шаг 8: Перевод решения
    while True:
        print_step_header(8, 13, "Перевод решения")
        if 'Solution' in df_bp_new.columns:
            unique_sols = get_unique_non_empty_values(df_bp_new['Solution'], 'Solution')

            if len(unique_sols) > 0:
                sol_translations = interactive_translation(unique_sols, "решений")
                df_bp_new['Solution (RUS)'] = df_bp_new['Solution'].map(sol_translations).fillna('-')
                df_bp_new = df_bp_new.drop(['Solution'], axis=1)
            else:
                print("  Колонка 'Solution': все значения равны '-' или пустые. Перевод не требуется.")
                df_bp_new['Solution (RUS)'] = '-'
                df_bp_new = df_bp_new.drop(['Solution'], axis=1)
            show_dataframe_preview(df_bp_new, "Перевод решения (после)",
                                  focus_columns=['Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Solution (RUS)'])

        if confirm_step("Перевод решения"):
            break

    # Шаг 9: Обработка цветов и Color Code
    while True:
        print_step_header(9, 13, "Обработка цветов и Color Code")

        if 'Color Name' in df_bp_new.columns:
            unique_colors = get_unique_non_empty_values(df_bp_new['Color Name'], 'Color Name')

            if len(unique_colors) > 0:
                print("  Выполняется перевод цветов...")
                color_translations = interactive_translation(unique_colors, "цветов")
                df_bp_new['Color Name (RUS)'] = df_bp_new['Color Name'].map(color_translations).fillna('-')
                df_bp_new = df_bp_new.drop(['Color Name'], axis=1)
            else:
                print("  Колонка 'Color Name': все значения равны '-' или пустые. Перевод не требуется.")
                df_bp_new['Color Name (RUS)'] = '-'
                df_bp_new = df_bp_new.drop(['Color Name'], axis=1)
        else:
            print("  Колонка 'Color Name' отсутствует. Создаём колонку 'Color Name (RUS)' со значениями '-'.")
            df_bp_new['Color Name (RUS)'] = '-'

        if 'Color Code' in df_bp_new.columns:
            unique_codes = get_unique_non_empty_values(df_bp_new['Color Code'], 'Color Code')

            if len(unique_codes) > 0:
                print("  Колонка 'Color Code' содержит данные. Исходные значения сохранены без изменений.")
            else:
                print("  Колонка 'Color Code' не содержит данных или все значения равны '-'. Заполняем '-'.")
                df_bp_new['Color Code'] = '-'
        else:
            print("  Колонка 'Color Code' отсутствует. Создаём колонку 'Color Code' со значениями '-'.")
            df_bp_new['Color Code'] = '-'

        show_dataframe_preview(df_bp_new, "Обработка цветов",
                              focus_columns=['Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Color Code', 'Color Name (RUS)'])

        if confirm_step("Обработка цветов и Color Code"):
            break

    # Шаг 10: Обработка рабочих центров
    while True:
        print_step_header(10, 13, "Обработка рабочих центров")
        if 'Workcenter Name' in df_bp_new.columns:
            df_bp_new['Workcenter Name'] = df_bp_new['Workcenter Name'].apply(extract_parentheses_content)

            unique_wc = get_unique_non_empty_values(df_bp_new['Workcenter Name'], 'Workcenter Name')

            if len(unique_wc) > 0:
                wc_translations = interactive_translation(unique_wc, "рабочих центров")
                df_bp_new['Workcenter Name'] = df_bp_new['Workcenter Name'].map(wc_translations).fillna('-')
            else:
                print("  Все значения рабочих центров равны '-' или пустые. Перевод не требуется.")
            show_dataframe_preview(df_bp_new, "Обработка рабочих центров",
                                  focus_columns=['Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Workcenter Name'])

        if confirm_step("Обработка рабочих центров"):
            break

    # Шаг 11: Проверка наличия в BOM
    print_step_header(11, 13, "Проверка наличия деталей в BOM")
    if df_bom is not None and 'BOM Product' in df_bp_new.columns and 'Part No.' in df_bp_new.columns:
        try:
            df_bp_new['Composite Key'] = df_bp_new['BOM Product'].astype(str) + '|' + df_bp_new['Part No.'].astype(str)
            df_bom['Composite Key'] = df_bom['Model'].astype(str) + '|' + df_bom['Part number'].astype(str)
            df_bp_new['Is in BOM'] = df_bp_new['Composite Key'].isin(df_bom['Composite Key'])
            df_bp_new = df_bp_new.drop(['Composite Key'], axis=1)

            if df_bp_new['Is in BOM'].dtype == bool:
                in_bom_count = df_bp_new['Is in BOM'].sum()
                print(f"  Результат: {in_bom_count} из {len(df_bp_new)} деталей найдены в BOM")
            else:
                print("  Результат: проверка выполнена")
        except KeyError as e:
            print(f"Ошибка: Отсутствует необходимая колонка в BOM файле: {e}")
            df_bp_new['Is in BOM'] = 'Error'
    else:
        df_bp_new['Is in BOM'] = 'Unknown'
        print("  Проверка не выполнена: отсутствуют необходимые колонки или BOM файл")

    show_dataframe_preview(df_bp_new, "Проверка наличия в BOM",
                          focus_columns=['Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Is in BOM'])
    if not confirm_step("Проверка наличия в BOM"):
        return process_bp_file(bp_filename, df_bom)

    # Шаг 12: Упорядочивание колонок
    print_step_header(12, 13, "Упорядочивание колонок")
    bp_columns_order = [
        'BP_No', 'In Stock', 'New Part Available Date', 'BOM Product', 'Change', 'Update Type',
        'Is in BOM', 'Part No.', 'Part Name (RUS)', 'Workcenter No.', 'Workcenter Name',
        'Quantity', 'Production Part Disposal', 'Interchangeable', 'Supplier Name (RUS)',
        'Change Description (RUS)', 'Solution (RUS)', 'Color Code', 'Color Name (RUS)',
    ]

    existing_cols = [col for col in bp_columns_order if col in df_bp_new.columns]
    missing_cols_in_order = set(bp_columns_order) - set(existing_cols)

    print(f"  Выбрано колонок для финального вывода: {len(existing_cols)} из {len(bp_columns_order)}")
    if missing_cols_in_order:
        print(f"  Отсутствуют в данных: {missing_cols_in_order}")

    df_bp_new = df_bp_new[existing_cols]
    show_dataframe_preview(df_bp_new, "Упорядочивание колонок (финальный результат)",
                          focus_columns=['BP_No', 'Change', 'Part No.', 'Part Name (RUS)', 'Is in BOM'])
    if not confirm_step("Упорядочивание колонок"):
        return process_bp_file(bp_filename, df_bom)

    # Шаг 13: Сохранение результата
    print_step_header(13, 13, "Сохранение результата")
    current_date = datetime.now().strftime('%Y-%m-%d')
    output_filename = f"{current_date}_{bp_number}_refactored.xlsx"
    save_excel_file(df_bp_new, output_filename)

    return output_filename


def main():
    """Главная функция программы"""
    clear_screen()

    print("""
        ╔══════════════════════════════════════════════════════════════╗
        ║               BREAKPOINT REFACTORING TOOL v1.0               ║
        ║               Пошаговая обработка Excel файлов               ║
        ╚══════════════════════════════════════════════════════════════╝
        """)

    print("Текущая директория:", os.getcwd())
    print("\nИнструкция:")
    print("   1. Программа предназначена для обработки технических изменений - Breakpoint (BP)")
    print("   2. Программа будет обрабатывать Excel файлы по одному")
    print("   3. Для каждого перевода нужно будет ввести русскую версию")
    print("   4. Нажмите Enter, чтобы пропустить перевод и оставить текст без изменений")
    print("   5. После каждого шага Вам предоставляется возможность проверить внесенные изменения:")
    print("     5.1. Если внесенные изменения корректны, нажмите Enter")
    print("     5.2. Если внесенные изменения некорректны, введите 'retry'")
    print("     5.3. 'retry' отменит внесенные изменения и Вы сможете исправить неточность")
    print("     5.4. После исправления нажмите Enter")
    print("   6. Нажмите клавиши Ctrl+C, чтобы прервать работу программы")

    wait_for_user()

    # Этап 1: Загрузка BOM файла
    print("\n" + "=" * 60)
    print("ЭТАП 1: Загрузка BOM файла")
    print("=" * 60)
    df_bom = load_excel_file('bom.xlsx', "BOM файл")
    if df_bom is None:
        print("\nКритическая ошибка: Не найден или не загружен файл bom.xlsx!")
        print("Убедитесь, что файл bom.xlsx находится в той же папке и имеет правильный формат.")
        wait_for_user()
        sys.exit(1)

    # Этап 2: Поиск BP файлов
    print("\n" + "=" * 60)
    print("ЭТАП 2: Поиск BP файлов")
    print("=" * 60)
    bp_files = find_bp_files()

    if not bp_files:
        print("Не найдено ни одного BP файла (формат: BP*.xlsx)")
        print("Убедитесь, что файлы начинаются с 'BP' и имеют расширение .xlsx")
        wait_for_user()
        sys.exit(1)

    print(f"Найдено BP файлов: {len(bp_files)}")
    for i, f in enumerate(bp_files, 1):
        print(f"{i}. {f}")

    wait_for_user()

    # Обработка каждого BP файла
    processed_files = []

    for i, bp_file in enumerate(bp_files, 1):
        print("\n" + "=" * 60)
        print(f"ОБРАБОТКА BP ФАЙЛА {i}/{len(bp_files)}: {bp_file}")
        print("=" * 60)

        print(f"\nТекущий файл: {bp_file}")
        print("Будут обработаны:")
        print("   - Перевод названий деталей")
        print("   - Перевод поставщиков")
        print("   - Фильтрация китайских символов")
        print("   - Перевод описаний и решений")
        print("   - Проверка наличия в BOM")

        wait_for_user("\nНажмите Enter для начала обработки этого файла...")

        result = process_bp_file(bp_file, df_bom)
        if result:
            processed_files.append(result)

        if i < len(bp_files):
            wait_for_user("\nФайл обработан. Нажмите Enter для перехода к следующему файлу...")

    # Итоги
    clear_screen()
    print("\n" + "=" * 60)
    print("ОБРАБОТКА ЗАВЕРШЕНА")
    print("=" * 60)
    print(f"\nОбработано файлов: {len(processed_files)}/{len(bp_files)}")
    print("\nСозданные файлы:")
    for f in processed_files:
        print(f"   - {f}")

    print("\nПрограмма завершила работу.")
    wait_for_user("\nНажмите Enter для выхода...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nПрограмма прервана пользователем")
        sys.exit(0)
    except BrokenPipeError:
        devnull = os.open(os.devnull, os.O_WRONLY)
        os.dup2(devnull, sys.stdout.fileno())
        sys.exit(1)

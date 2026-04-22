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
        print(f"Размер: {df.shape[0]} строк × {df.shape[1]} столбцов")
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
        # Непредвиденная ошибка - логируем детали для отладки
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

    print("\nВводите переводы для каждого значения.")
    print("Если оставить пустым - значение будет пропущено (заменено на '-')")
    print("Введите 'skip' чтобы пропустить все оставшиеся")
    print("Введите 'done' чтобы завершить ввод")

    for i, value in enumerate(data, 1):
        if pd.isna(value) or value == '':
            translations[value] = '-'
            continue

        print(f"\n   [{i}/{len(data)}] Оригинал: {value}")
        try:
            user_input = input(
                "Перевод (Enter - пропустить, 'skip' - все пропустить, 'done' - закончить): "
            ).strip()
        except EOFError:
            print("\nОбнаружен конец ввода. Пропускаем оставшиеся переводы.")
            translations[value] = '-'
            for remaining in data[i:]:
                translations[remaining] = '-'
            break
        except KeyboardInterrupt:
            print("\nВвод прерван. Пропускаем оставшиеся переводы.")
            translations[value] = '-'
            for remaining in data[i:]:
                translations[remaining] = '-'
            break

        if user_input.lower() == 'done':
            for remaining in data[i - 1:]:
                translations[remaining] = '-'
            break
        elif user_input.lower() == 'skip':
            translations[value] = '-'
            for remaining in data[i:]:
                translations[remaining] = '-'
            break
        elif user_input == '':
            translations[value] = '-'
        else:
            translations[value] = user_input

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


def process_bp_file(bp_filename, df_bom):
    """
    Обработка одного BP файла
    """
    print(f"\nОбработка файла: {bp_filename}")

    bp_number = bp_filename.replace('.xlsx', '')

    # Шаг 1: Загрузка BP файла
    df_bp = load_excel_file(bp_filename, "BP файла")
    if df_bp is None:
        return None

    # Шаг 2: Выбор нужных колонок
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

    # Шаг 3: Перевод названий деталей
    if 'Part Name(CHN)' in df_bp_new.columns:
        print("\nШаг 3: Перевод названий деталей")
        part_names = df_bp_new['Part Name(CHN)'].dropna().unique()

        if len(part_names) > 0:
            print(f"Найдено {len(part_names)} уникальных названий деталей")
            part_translations = interactive_translation(part_names, "названий деталей")
            df_bp_new['Part Name (RUS)'] = df_bp_new['Part Name(CHN)'].map(part_translations).fillna('-')
            df_bp_new = df_bp_new.drop(columns=['Part Name(CHN)'], axis=1)

    # Шаг 4: Перевод поставщиков
    if 'Supplier Name' in df_bp_new.columns:
        print("\nШаг 4: Перевод названий поставщиков")
        suppliers = df_bp_new['Supplier Name'].dropna().unique()

        if len(suppliers) > 0:
            supplier_translations = interactive_translation(suppliers, "поставщиков")
            df_bp_new['Supplier Name (RUS)'] = df_bp_new['Supplier Name'].map(supplier_translations).fillna('-')
            df_bp_new = df_bp_new.drop(columns=['Supplier Name'], axis=1)

    # Шаг 5: Фильтрация китайских символов
    print("\nШаг 5: Фильтрация китайских символов")
    if 'Change Description' in df_bp_new.columns:
        df_bp_new['Change Description'] = df_bp_new['Change Description'].apply(filter_chinese_lines)

    if 'Solution' in df_bp_new.columns:
        df_bp_new['Solution'] = df_bp_new['Solution'].apply(filter_chinese_lines)

    # Шаг 6: Перевод описаний изменений
    if 'Change Description' in df_bp_new.columns:
        print("\nШаг 6: Перевод описаний изменений")
        descriptions = df_bp_new['Change Description'].dropna().unique()
        if len(descriptions) > 0:
            desc_translations = interactive_translation(descriptions, "описаний изменений")
            df_bp_new['Change Description (RUS)'] = df_bp_new['Change Description'].map(desc_translations).fillna('-')
            df_bp_new = df_bp_new.drop(columns=['Change Description'], axis=1)

    # Шаг 7: Перевод решений
    if 'Solution' in df_bp_new.columns:
        print("\nШаг 7: Перевод решений")
        solutions = df_bp_new['Solution'].dropna().unique()
        if len(solutions) > 0:
            sol_translations = interactive_translation(solutions, "решений")
            df_bp_new['Solution (RUS)'] = df_bp_new['Solution'].map(sol_translations).fillna('-')
            df_bp_new = df_bp_new.drop(columns=['Solution'], axis=1)

    # Шаг 8: Перевод цветов
    if 'Color Name' in df_bp_new.columns:
        print("\nШаг 8: Перевод названий цветов")
        colors = df_bp_new['Color Name'].dropna().unique()
        if len(colors) > 0:
            color_translations = interactive_translation(colors, "цветов")
            df_bp_new['Color Name (RUS)'] = df_bp_new['Color Name'].map(color_translations).fillna('-')
            df_bp_new = df_bp_new.drop(columns=['Color Name'], axis=1)

    # Шаг 9: Обработка рабочих центров
    if 'Workcenter Name' in df_bp_new.columns:
        print("\nШаг 9: Обработка рабочих центров")
        df_bp_new['Workcenter Name'] = df_bp_new['Workcenter Name'].apply(extract_parentheses_content)

        workcenters = df_bp_new['Workcenter Name'].dropna().unique()
        if len(workcenters) > 0 and not (len(workcenters) == 1 and workcenters[0] == '-'):
            wc_translations = interactive_translation(workcenters, "рабочих центров")
            df_bp_new['Workcenter Name'] = df_bp_new['Workcenter Name'].map(wc_translations).fillna('-')

    # Шаг 10: Проверка наличия в BOM
    print("\nШаг 10: Проверка наличия деталей в BOM")
    if df_bom is not None and 'BOM Product' in df_bp_new.columns and 'Part No.' in df_bp_new.columns:
        try:
            df_bp_new['Composite Key'] = df_bp_new['BOM Product'].astype(str) + '|' + df_bp_new['Part No.'].astype(str)
            df_bom['Composite Key'] = df_bom['Model'].astype(str) + '|' + df_bom['Part number'].astype(str)
            df_bp_new['Is in BOM'] = df_bp_new['Composite Key'].isin(df_bom['Composite Key'])
            df_bp_new = df_bp_new.drop(columns=['Composite Key'], axis=1)
        except KeyError as e:
            print(f"Ошибка: Отсутствует необходимая колонка в BOM файле: {e}")
            df_bp_new['Is in BOM'] = 'Error'
    else:
        df_bp_new['Is in BOM'] = 'Unknown'

    # Шаг 11: Упорядочивание колонок
    bp_columns_order = [
        'BP_No', 'In Stock', 'New Part Available Date', 'BOM Product', 'Change', 'Update Type',
        'Is in BOM', 'Part No.', 'Part Name (RUS)', 'Workcenter No.', 'Workcenter Name',
        'Quantity', 'Production Part Disposal', 'Interchangeable', 'Supplier Name (RUS)',
        'Change Description (RUS)', 'Solution (RUS)', 'Color Code', 'Color Name (RUS)',
    ]

    existing_cols = [col for col in bp_columns_order if col in df_bp_new.columns]
    df_bp_new = df_bp_new[existing_cols]

    # Шаг 12: Сохранение результата
    current_date = datetime.now().strftime('%Y-%m-%d')
    output_filename = f"{current_date}_{bp_number}_refactored.xlsx"
    save_excel_file(df_bp_new, output_filename)

    return output_filename


def main():
    """Главная функция программы"""
    clear_screen()

    print("""
        ╔══════════════════════════════════════════════════════════════╗
        ║                   BP REFACTORING TOOL v1.0                   ║
        ║           Пошаговая обработка Excel файлов BP                ║
        ╚══════════════════════════════════════════════════════════════╝
        """)

    print("Текущая директория:", os.getcwd())
    print("\nИнструкция:")
    print("   1. Программа будет обрабатывать BP файлы по одному")
    print("   2. Для каждого перевода нужно будет ввести русскую версию")
    print("   3. Enter - пропустить текущий перевод, 'skip' - пропустить все оставшиеся, 'done' - завершить ввод переводов для текущей категории")

    wait_for_user()

    # Шаг 0: Загрузка BOM файла
    print_step_header(0, 5, "Загрузка BOM файла")
    df_bom = load_excel_file('bom.xlsx', "BOM файла")
    if df_bom is None:
        print("\nКритическая ошибка: Не найден или не загружен файл bom.xlsx!")
        print("Убедитесь, что файл bom.xlsx находится в той же папке и имеет правильный формат.")
        wait_for_user()
        sys.exit(1)

    # Поиск BP файлов
    print_step_header(1, 5, "Поиск BP файлов")
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
        print_step_header(1 + i, 5 + len(bp_files), f"Обработка файла {i}/{len(bp_files)}")

        print(f"\nТекущий файл: {bp_file}")
        print("Будут обработаны:")
        print("   - Перевод названий деталей")
        print("   - Перевод поставщиков")
        print("   - Фильтрация китайских символов")
        print("   - Перевод описаний")
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
        # При работе с пайпами в Windows может возникать
        devnull = os.open(os.devnull, os.O_WRONLY)
        os.dup2(devnull, sys.stdout.fileno())
        sys.exit(1)

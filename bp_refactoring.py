#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
"""
BP Refactoring Tool - Пошаговая обработка Excel файлов

Этот модуль отвечает за загрузку и обработку индивидуальных Excel файлов

Breakpoint (BP). Он выполняет 18 последовательных шагов обработки:
    1. Загрузка BP файла
    2. Выбор нужных колонок
    3. Выбор статуса технического изменения
    4. Заполнение пустых значений
    5. Ввод количества деталей в SS
    6. Перевод названий деталей
    7. Поиск официальных названий поставщиков
    8. Ввод статуса локализации поставщиков
    9. Фильтрация китайских символов
    10. Перевод описания к изменению
    11. Перевод решения к изменению
    12. Обработка цветов и Color Code
    13. Обработка рабочих центров
    14. Перевод требований по утилизации старых деталей
    15. Перевод требований по взаимозаменяемости
    16. Проверка наличия деталей в BOM
    17. Упорядочивание колонок
    18. Сохранение результата

Каждый шаг включает:
    - Интерактивное взаимодействие с пользователем для ввода переводов
    - Возможность отменить изменения и повторить шаг (режим retry)
    - Сохранение состояния перед каждым шагом для возможности восстановления

Модуль также содержит логику определения новых BP через сравнение
с existing breakpoint_data и bp_list_2025-2026.xlsx.

Использование:
    from bp_refactoring import main as refactoring_main

Версия: 1.0
Совместимость: Python 3.12.3+, Pandas 3.0.2+, OpenPyXL 3.1.5+
Поддержка: PLD Engineering Center
Дата создания: 2026-05-07
Лицензия: MIT
Статус: Production
"""
import io
import os
import re
import sys
import warnings
from datetime import datetime
from typing import Optional

import pandas as pd

from pandas.errors import EmptyDataError, ParserError
from openpyxl.utils.exceptions import InvalidFileException

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Для Windows консоли
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


def safe_str_convert(
        value,
        default=''
    ) -> str:
    """
    Безопасно преобразует значение в строку, обрабатывая None и ошибки.

    Аргументы:
        value: Значение для преобразования.
        default (str): Значение по умолчанию при ошибке. По умолчанию ''.

    Возвращается:
        str: Преобразованная строка или значение по умолчанию.
    """
    if value is None:
        return default
    try:
        return str(value).strip()
    except (ValueError, TypeError):
        return default


def safe_date_convert(
        value,
        default=None,
        pattern: str = '%Y-%m-%d'
    ) -> Optional[datetime]:
    """
    Безопасно преобразует значение в дату, обрабатывая различные форматы.

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


def wait_for_user(prompt="\nНажмите Enter для продолжения..."):
    """
    Ожидает нажатия клавиши Enter от пользователя.

    Обрабатывает Ctrl+C и EOF для корректного завершения программы.

    Аргументы:
        prompt (str): Текст приглашения к вводу.
    """
    try:
        input(prompt)
    except KeyboardInterrupt:
        print("\n\nПрограмма прервана пользователем (Ctrl+C)")
        sys.exit(0)
    except EOFError:
        print("\n\nОбнаружен конец ввода. Программа завершена.")
        sys.exit(0)


def print_step_header(step_num, total_steps, description):
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


def save_state_before_step(df):
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


def restore_state(saved_df, step_name):
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


def load_excel_file(filename, description="файл"):
    """
    Загружает Excel файл с полной обработкой ошибок и проверкой существования.

    Обрабатываемые ошибки:
        - FileNotFoundError
        - PermissionError (файл открыт в Excel)
        - EmptyDataError (пустой файл)
        - ParserError (повреждённый формат)
        - InvalidFileException (некорректный Excel файл)
        - ValueError (неопределённый формат)
        - Общие исключения

    Аргументы:
        filename (str): Имя файла для загрузки.
        description (str): Описание файла для вывода сообщений.

    Возвращается:
        pd.DataFrame: Загруженный DataFrame или None при ошибке.
    """
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


def show_dataframe_preview(df, step_name, max_rows=10, focus_columns=None, max_colwidth=40):
    """
    Отображает первые строки DataFrame для визуального контроля результатов шага.

    Аргументы:
        df (pd.DataFrame): DataFrame для отображения.
        step_name (str): Имя шага для вывода в заголовке.
        max_rows (int): Максимальное количество строк для отображения. По умолчанию 5.
        focus_columns (list, optional): Список колонок для отображения.
                                        Если None, показываются первые 5 колонок.
        max_colwidth (int): Максимальная ширина содержимого колонки в символах.
                            По умолчанию 40.
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
    """
    Фильтрует китайские иероглифы из текста.

    Для многострочного текста собирает строки с китайскими иероглифами.
    Для однострочного текста обрезает всё после последнего китайского иероглифа.

    Аргументы:
        text: Текст для фильтрации (может быть не строкой).

    Возвращается:
        str: Отфильтрованный текст или исходное значение, если не строка.
    """
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


def fill_empty_values_with_dash(df, columns):
    """
    Заполняет пустые значения в указанных колонках символом '-'.

    Аргументы:
        df (pd.DataFrame): DataFrame для обработки.
        columns (list): Список названий колонок для обработки.

    Возвращается:
        list: Список колонок, в которых были выполнены замены.

    Примечание:
        Пустыми считаются значения: None, NaN, пустая строка, 'nan', 'None'.
    """
    columns_with_replacements = []

    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str)

            empty_mask = df[col].isna() | (df[col].str.strip() == '') | (df[col].str.strip() == 'nan') | (df[col].str.strip() == 'None')
            empty_count = empty_mask.sum()

            if empty_count > 0:
                print(f"  Колонка '{col}': ячейки без данных - {empty_count}. Заполнено '-'.")
                df.loc[empty_mask, col] = '-'
                columns_with_replacements.append(col)
            else:
                print(f"  Колонка '{col}': все данные заполнены. Сохранено без изменений.")
        else:
            print(f"  Колонка '{col}': отсутствует в данных.")

    return columns_with_replacements


def get_unique_non_empty_values(series, column_name):
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


def interactive_translation(data, field_name, examples=None):
    """
    Интерактивный ввод переводов для списка уникальных значений.

    Пользователь может:
    - Ввести перевод → значение будет заменено
    - Нажать Enter без ввода → значение останется оригинальным

    Аргументы:
        data (list): Список уникальных значений для перевода.
        field_name (str): Название поля для вывода (например, "названий деталей").
        examples (dict, optional): Словарь примеров переводов для отображения.

    Возвращается:
        dict: Словарь соответствий {оригинал: перевод}.

    Примечание:
        Функция отображает все значения перед началом ввода.
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

    # Выводим все оригинальные значения перед началом ввода
    print("\nСписок всех уникальных значений:")
    print("-" * 60)
    for i, value in enumerate(data, 1):
        print(f"  {i}. {value}")
    print("-" * 60)

    print("\nИнструкция:")
    print("  • Введите перевод и нажмите Enter → оригинальный текст будет заменён на перевод")
    print("  • Нажмите Enter без перевода → текст останется оригинальным (без изменений)")

    for i, value in enumerate(data, 1):
        if pd.isna(value) or value == '':
            translations[value] = value
            continue

        print(f"\n[{i}/{len(data)}] Оригинал: {value}")
        try:
            user_input = input(
                "Введите перевод (или просто Enter чтобы оставить оригинал): "
            ).strip()
        except KeyboardInterrupt:
            print("\n\nПрограмма прервана пользователем (Ctrl+C)")
            sys.exit(0)
        except EOFError:
            print("\n\nКонец ввода. Программа завершена.")
            sys.exit(0)

        if user_input == '':
            translations[value] = value
            print(f"  → Оставляем оригинал: {value}")
        else:
            translations[value] = user_input
            print(f"  → Заменяем на: {user_input}")

    return translations


def find_bp_files():
    """
    Ищет BP файлы в текущей папке.

    Файлы должны соответствовать шаблону: BP*.xlsx
    Временные файлы (~$BP*.xlsx) игнорируются.

    Возвращается:
        list: Отсортированный список найденных BP файлов.
    """
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


def find_latest_breakpoint_file(file_prefix: str = 'breakpoint_data') -> Optional[str]:
    """
    Находит файл breakpoint_data с самой поздней датой в имени.

    Аргументы:
        file_prefix (str): Префикс имени файла для поиска.

    Возвращается:
        Optional[str]: Имя самого свежего файла или None, если файлы не найдены.

    Примечание:
        bp_main.find_latest_breakpoint_file - аналогичная функция в главном модуле
    """
    pattern = rf"[0-9]{{4}}-[0-9]{{2}}-[0-9]{{2}}_{file_prefix}\.xlsx"

    matching_files = []

    for filename in os.listdir('.'):
        if re.match(pattern, filename):
            # Извлекаем дату из имени файла
            date_str = filename[:10]  # первые 10 символов = YYYY-MM-DD
            try:
                file_date = datetime.strptime(date_str, '%Y-%m-%d')
                matching_files.append((file_date, filename))
            except ValueError:
                continue

    if not matching_files:
        return None

    # Сортируем по дате и возвращаем самый свежий
    matching_files.sort(key=lambda x: x[0], reverse=True)
    return matching_files[0][1]


def load_latest_breakpoint_data(file_prefix: str = 'breakpoint_data') -> Optional[pd.DataFrame]:
    """
    Загружает самый свежий файл breakpoint_data с датой в имени.

    Аргументы:
        file_prefix (str): Префикс имени файла для поиска.

    Возвращается:
        Optional[pd.DataFrame]: DataFrame с данными или None, если файлы не найдены.
    """
    latest_file = find_latest_breakpoint_file(file_prefix)

    if latest_file is None:
        print("  Внимание: Не найдено ни одного файла breakpoint_data с историей")
        return None

    try:
        # Загружаем с указанием строки заголовка (3-я строка = header=2)
        df = pd.read_excel(latest_file, header=2)
        print(f"  Загружен последний файл: {latest_file}")
        print(f"  В нём {len(df)} строк")
        return df
    except FileNotFoundError:
        print(f"  Ошибка: Файл '{latest_file}' не найден")
        return None
    except PermissionError:
        print(f"  Ошибка: Нет прав для чтения файла '{latest_file}'")
        print("  Закройте файл, если он открыт в Excel, и попробуйте снова.")
        return None
    except (EmptyDataError, ParserError) as e:
        print(f"  Ошибка: Файл '{latest_file}' повреждён или имеет неверный формат: {e}")
        return None
    except InvalidFileException:
        print(f"  Ошибка: Файл '{latest_file}' не является корректным Excel файлом")
        return None
    except ValueError as e:
        if "Excel file format cannot be determined" in str(e):
            print(f"  Ошибка: Не удалось определить формат файла '{latest_file}'")
            print("  Убедитесь, что файл имеет расширение .xlsx или .xls")
        else:
            print(f"  Ошибка при загрузке файла '{latest_file}': {e}")
        return None
    except Exception as e:
        # Непредвиденная ошибка
        print(f"  НЕПРЕДВИДЕННАЯ ОШИБКА при загрузке файла '{latest_file}': {e}")
        print(f"  Тип ошибки: {type(e).__name__}")
        return None


def new_bp_check():
    """
    Проверяет наличие новых BP для скачивания из системы G-BOM.

    Алгоритм:
    1. Загружает bp_list_2025-2026.xlsx
    2. Загружает последний файл breakpoint_data.xlsx
    3. Применяет фильтры к списку BP:
       - IsUnBomBP = "No"
       - Status != 'Closed'
       - Date >= 01.04.2026 или пустая дата
       - Part Name (E) не содержит "SOFTWARE"
    4. Находит разницу между BP в отфильтрованном списке и уже обработанными

    Возвращается:
        set: Множество номеров BP, которые нужно скачать, или None при ошибке.

    Требования:
        - Файл bp_list_2025-2026.xlsx должен существовать в текущей папке
        - Доступ к файлу для чтения
    """

    # Проверка наличия файлов перед загрузкой
    required_files = ['bp_list_2025-2026.xlsx']
    missing_files = []

    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)

    if missing_files:
        print("\n  ОШИБКА: Отсутствуют обязательные файлы:")
        for file in missing_files:
            print(f"    - {file}")
        print(f"\n  Текущая директория: {os.getcwd()}")
        print("  Убедитесь, что все файлы находятся в текущей папке.")
        return None

    # 1) Загрузка bp_list_2025-2026.xlsx с обработкой ошибок
    print("\n  Загрузка файла bp_list_2025-2026.xlsx...")
    df_source_bp_list = None

    try:
        # Проверка прав доступа к файлу
        if not os.access('bp_list_2025-2026.xlsx', os.R_OK):
            print("  ОШИБКА: Нет прав на чтение файла bp_list_2025-2026.xlsx!")
            print("  Закройте файл, если он открыт в Excel, и попробуйте снова.")
            return None

        df_source_bp_list = pd.read_excel('bp_list_2025-2026.xlsx')
        print("    Файл 'bp_list_2025-2026.xlsx' загружен успешно!")
        print(f"    Размер: {df_source_bp_list.shape[0]} строк × {df_source_bp_list.shape[1]} колонок")

    except PermissionError:
        print("  ОШИБКА: Нет прав доступа к файлу bp_list_2025-2026.xlsx!")
        print("  Закройте файл, если он открыт в других программах, и попробуйте снова.")
        return None
    except FileNotFoundError:
        print("  ОШИБКА: Файл bp_list_2025-2026.xlsx не найден!")
        return None
    except EmptyDataError:
        print("  ОШИБКА: Файл bp_list_2025-2026.xlsx пуст!")
        return None
    except ParserError as e:
        print(f"  ОШИБКА: Не удалось разобрать файл bp_list_2025-2026.xlsx: {e}")
        return None
    except InvalidFileException:
        print("  ОШИБКА: Файл bp_list_2025-2026.xlsx не является корректным Excel файлом!")
        return None
    except ValueError as e:
        if "Excel file format cannot be determined" in str(e):
            print("  ОШИБКА: Не удалось определить формат файла bp_list_2025-2026.xlsx!")
            print("  Убедитесь, что файл имеет расширение .xlsx или .xls")
        else:
            print(f"  ОШИБКА при загрузке bp_list_2025-2026.xlsx: {e}")
        return None
    except Exception as e:
        print(f"  НЕПРЕДВИДЕННАЯ ОШИБКА при загрузке bp_list_2025-2026.xlsx: {e}")
        print(f"  Тип ошибки: {type(e).__name__}")
        return None

    # 2) Загрузка самого свежего breakpoint_data.xlsx с датой в имени
    print("\n  Загрузка последнего файла breakpoint_data...")
    df_source_breakpoint_data = load_latest_breakpoint_data()

    if df_source_breakpoint_data is None or df_source_breakpoint_data.empty:
        print("  ВНИМАНИЕ: Не найдено файлов breakpoint_data с историей")
        print("  Все BP из списка будут считаться новыми.")
        set_bp_no = set()
    else:
        # Проверяем, что колонка BP_No действительно существует
        if 'BP_No' not in df_source_breakpoint_data.columns:
            print("  ОШИБКА: Колонка 'BP_No' не найдена в загруженном файле!")
            print(f"  Доступные колонки: {list(df_source_breakpoint_data.columns)[:10]}...")
            return None

        # Очистка данных: удаляем пустые значения и NaN
        bp_no_series = df_source_breakpoint_data['BP_No'].dropna()
        if bp_no_series.empty:
            print("  ВНИМАНИЕ: Колонка 'BP_No' не содержит значений!")
            set_bp_no = set()
        else:
            set_bp_no = set(bp_no_series.astype(str).str.strip())
            set_bp_no = {bp for bp in set_bp_no if bp and bp != 'nan' and bp != 'None'}

        print(f"  Уникальных BP в последнем файле: {len(set_bp_no)}")

    # 3) Применение фильтров к df_source_bp_list (как в test_bp_check.py)
    print("\n  Применение фильтров к bp_list_2025-2026.xlsx...")

    # Удаляем полностью пустые строки
    df_filtered = df_source_bp_list.dropna(how='all').copy()

    # Проверка наличия необходимых колонок
    required_columns = ['IsUnBomBP', 'Status', 'Date', 'Part Name (E)', 'BP']
    missing_columns = [col for col in required_columns if col not in df_filtered.columns]

    if missing_columns:
        print("\n  ОШИБКА: В файле bp_list_2025-2026.xlsx отсутствуют необходимые колонки:")
        for col in missing_columns:
            print(f"    - {col}")
        print(f"\n  Доступные колонки: {list(df_filtered.columns)}")
        return None

    # Преобразование колонок (как в test_bp_check.py)
    print("  Преобразование колонок...")
    df_filtered['IsUnBomBP'] = df_filtered['IsUnBomBP'].apply(safe_str_convert)
    df_filtered['Status'] = df_filtered['Status'].apply(safe_str_convert)
    df_filtered['Part Name (E)'] = df_filtered['Part Name (E)'].apply(safe_str_convert)

    df_filtered['Date'] = df_filtered['Date'].apply(safe_date_convert)

    # Фильтр 1: IsUnBomBP = "No"
    initial_count = len(df_filtered)
    mask1 = df_filtered['IsUnBomBP'] == 'No'
    df_filtered = df_filtered[mask1]
    print(f"    - Фильтр IsUnBomBP = 'No': {initial_count} → {len(df_filtered)} строк")

    # Фильтр 2: Status != 'Closed'
    initial_count = len(df_filtered)
    mask2 = df_filtered['Status'] != 'Closed'
    df_filtered = df_filtered[mask2]
    print(f"    - Фильтр Status != 'Closed': {initial_count} → {len(df_filtered)} строк")

    # Фильтр 3: Date >= 01.04.2026 ИЛИ дата пустая
    initial_count = len(df_filtered)
    cutoff_date = pd.to_datetime('2026-04-01')
    mask3 = (df_filtered['Date'] >= cutoff_date) | (df_filtered['Date'].isna())
    df_filtered = df_filtered[mask3]
    print(f"    - Фильтр Date >= 01.04.2026 или пустая дата: {initial_count} → {len(df_filtered)} строк")

    # Фильтр 4: Part Name (E) НЕ содержит слово SOFTWARE
    initial_count = len(df_filtered)
    software_mask = df_filtered['Part Name (E)'].str.lower().str.contains('software', na=False)
    df_filtered = df_filtered[~software_mask]
    print(f"    - Фильтр Part Name (E) НЕ содержит 'SOFTWARE': {initial_count} → {len(df_filtered)} строк")

    print(f"\n  После применения всех фильтров: {len(df_filtered)} строк")

    # 4) Создание set по колонке "BP" из отфильтрованного DataFrame
    try:
        if 'BP' not in df_filtered.columns:
            print("  ОШИБКА: Колонка 'BP' не найдена в отфильтрованных данных!")
            return None

        # Очистка данных: удаляем пустые значения и NaN
        bp_series = df_filtered['BP'].dropna()
        if bp_series.empty:
            print("  ВНИМАНИЕ: Колонка 'BP' не содержит значений после фильтрации!")
            set_bp = set()
        else:
            set_bp = set(bp_series.astype(str).str.strip())
            set_bp = {bp for bp in set_bp if bp and bp != 'nan' and bp != 'None'}

        print(f"\n  Уникальных BP из bp_list (после фильтров): {len(set_bp)}")
        if len(set_bp) > 0 and len(set_bp) <= 20:
            print(f"    Список: {sorted(set_bp)}")
        elif len(set_bp) > 20:
            print(f"    (Слишком много BP для отображения: {len(set_bp)} шт.)")

    except KeyError as e:
        print(f"  ОШИБКА: Колонка 'BP' не найдена при создании set: {e}")
        return None
    except AttributeError as e:
        print(f"  ОШИБКА: Не удалось обработать данные колонки 'BP': {e}")
        return None
    except ValueError as e:
        print(f"  ОШИБКА: Неверный формат данных в колонке 'BP': {e}")
        return None
    except Exception as e:
        print(f"  НЕПРЕДВИДЕННАЯ ОШИБКА при создании set BP: {e}")
        print(f"    Тип ошибки: {type(e).__name__}")
        return None

    # 5) Нахождение разницы (BP, которых нет в breakpoint_data)
    try:
        new_bp_set = set_bp - set_bp_no

        print("\n" + "=" * 60)
        print("  РЕЗУЛЬТАТ ПРОВЕРКИ НОВЫХ BP:")
        print("=" * 60)
        print(f"  Всего BP в списке (после фильтров): {len(set_bp)}")
        print(f"  BP уже в breakpoint_data: {len(set_bp & set_bp_no)}")
        print(f"  НОВЫЕ BP (требуют скачивания): {len(new_bp_set)}")

        if len(new_bp_set) > 0:
            print("\n  Список новых BP для скачивания из G-BOM:")
            sorted_new_bp = sorted(new_bp_set)
            for i, bp in enumerate(sorted_new_bp, 1):
                print(f"    {i}. {bp}")
        else:
            print("\n  Нет новых BP для скачивания.")

        return new_bp_set

    except TypeError as e:
        print(f"\n  ОШИБКА ТИПОВ при вычислении разницы BP: {e}")
        print("  Проверьте, что set_bp и set_bp_no являются множествами (set)")
        return None
    except NameError as e:
        print(f"\n  ОШИБКА: Переменные не определены - {e}")
        print("  Проверьте, что set_bp и set_bp_no были успешно созданы")
        return None
    except Exception as e:
        print(f"\n  КРИТИЧЕСКАЯ ОШИБКА при вычислении разницы BP: {e}")
        print(f"  Тип ошибки: {type(e).__name__}")
        return None


def confirm_step(step_name, df_bp_new, saved_state):
    """
    Запрашивает у пользователя подтверждение после выполнения шага.

    Поддерживает:
    - Enter: подтверждение, сохранение нового состояния
    - 'retry': отмена изменений и повтор шага

    Аргументы:
        step_name (str): Имя шага для вывода сообщений.
        df_bp_new (pd.DataFrame): Текущий DataFrame после выполнения шага.
        saved_state (pd.DataFrame): Сохранённое состояние перед шагом.

    Возвращается:
        tuple: (continue_flag, df, new_saved_state)
            - continue_flag (bool): True если нужно продолжить, False если retry
            - df (pd.DataFrame): Текущий или восстановленный DataFrame
            - new_saved_state (pd.DataFrame): Новое сохранённое состояние
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
            # Возвращаем False и восстановленный DataFrame, состояние не меняем
            return False, restored_df, saved_state
        else:
            return False, df_bp_new, saved_state
    else:
        print("Продолжаем...\n")
        # Сохраняем новое состояние после успешного шага
        new_saved_state = save_state_before_step(df_bp_new)
        return True, df_bp_new, new_saved_state


def get_bp_status(bp_number):
    """
    Запрашивает у пользователя статус для BP файла.

    Доступные варианты:
    1. Согласован/Approved
    2. Опубликован/Published
    3. Закрыт/Closed
    4. Другое (ввод вручную)

    Аргументы:
        bp_number (str): Номер BP для вывода в сообщении.

    Возвращается:
        str: Выбранный статус (с возможным переводом строки для двустрочного формата).
    """
    print(f"\n  Для BP файла {bp_number} укажите статус обработки:")
    print("  Доступные варианты:")
    print("    1 - Согласован\n\t\tApproved")
    print("    2 - Опубликован\n\t\tPublished")
    print("    3 - Закрыт\n\t\tClosed")
    print("    4 - Другое (ввести вручную)")

    while True:
        try:
            choice = input("  Ваш выбор (1-4, или Enter для 'Согласован/Approved'): ").strip()

            if choice == '' or choice == '1':
                return "Согласован\nApproved"
            elif choice == '2':
                return "Опубликован\nPublished"
            elif choice == '3':
                return "Закрыт\nClosed"
            elif choice == '4':
                custom_status = input("  Введите свой статус: ").strip()
                return custom_status if custom_status else "Согласован\nApproved"
            else:
                print("  Неверный выбор. Пожалуйста, введите число от 1 до 4.")
        except KeyboardInterrupt:
            print("\n\nПрограмма прервана пользователем (Ctrl+C)")
            sys.exit(0)


def classify_row_for_quantity_ss(row: pd.Series) -> str:
    """
    Классифицирует строку для определения, нужно ли вводить количество в SS.
    Количество в SS вводится ТОЛЬКО для старых деталей (Delete).

    Алгоритм:
    1. Приоритет 1: колонка 'Change'
        - 'Before Change' → Before
        - 'After Change' → After
    2. Приоритет 2: если Change пустой, то по 'Update Type':
        - 'Delete' → Before (старая деталь, была на складе)
        - 'Add' → After (новая деталь)
        - 'Replace' → After (новая деталь)
        - 'Update' → After (новая деталь)

    Аргументы:
        row (pd.Series): Строка DataFrame для классификации.

    Возвращается:
        str: 'Before' (требует ввода количества) или 'After' (не требует).
    """
    change_val = safe_str_convert(row.get('Change', ''))

    # Приоритет 1: явное указание в колонке Change
    if change_val == 'Before Change':
        return 'Before'
    elif change_val == 'After Change':
        return 'After'

    # Приоритет 2: определяем по Update Type
    update_type = safe_str_convert(row.get('Update Type', ''))

    # Только Delete детали требуют ввода количества в SS
    if update_type == 'Delete':
        return 'Before'
    else:
        # Add, Replace, Update — это новые детали
        return 'After'


def get_quantity_in_ss(df_bp_new, bp_number):
    """
    Интерактивный ввод количества деталей в SS для СТАРЫХ ДЕТАЛЕЙ (Delete).

    Пользователь:
    - Вводит номер детали
    - Вводит количество (целое число)
    - Можно ввести несколько деталей последовательно
    - Enter без ввода номера устанавливает 0 для всех оставшихся деталей

    Аргументы:
        df_bp_new (pd.DataFrame): DataFrame с данными BP.
        bp_number (str): Номер BP для вывода.

    Возвращается:
        dict: Словарь {Part No.: количество} для деталей Delete.
    """
    # Определяем, какие строки требуют ввода количества в SS (только Delete = Before)
    df_bp_new['_need_quantity'] = df_bp_new.apply(classify_row_for_quantity_ss, axis=1)

    # Фильтруем только детали, для которых нужно вводить количество (Before = Delete)
    before_df = df_bp_new[df_bp_new['_need_quantity'] == 'Before'].copy()

    # Получаем уникальные детали Before (без дубликатов)
    unique_parts = before_df[['Part No.']].drop_duplicates(subset=['Part No.'])

    # Фильтруем пустые значения
    valid_parts = []
    for _, row in unique_parts.iterrows():
        part_no = row.get('Part No.', '')
        if not pd.isna(part_no) and str(part_no).strip() != '' and str(part_no).strip() != '-':
            part_no_str = str(part_no).strip()
            valid_parts.append(part_no_str)

    total_parts = len(valid_parts)

    # Выводим статистику по классификации
    total_before = len(before_df)
    total_after = len(df_bp_new) - total_before
    print(f"\n  Анализ деталей для BP {bp_number}:")
    print(f"    - Деталей Delete (старые, требуют ввод количества): {total_before} записей, {len(valid_parts)} уникальных")
    print(f"    - Деталей Add/Replace/Update (новые, количество = 0): {total_after} записей")

    if total_parts == 0:
        print("\n  ВНИМАНИЕ: Не найдено деталей Delete для ввода количества в SS!")
        return {}

    print(f"\n  Ввод количества деталей в SS для BP {bp_number}")
    print("-" * 60)
    print("  СПИСОК УНИКАЛЬНЫХ ДЕТАЛЕЙ (DELETE - старые детали):")
    for i, part_no in enumerate(valid_parts, 1):
        print(f"    {i}. {part_no}")
    print("-" * 60)

    print("\n  ИНСТРУКЦИЯ:")
    print("    1. Введите номер детали, для которой хотите указать количество")
    print("    2. Затем введите количество (целое число)")
    print("    3. Можно ввести несколько деталей последовательно")
    print("    4. Нажмите Enter без ввода номера детали, чтобы установить 0 для ВСЕХ оставшихся деталей")
    print("-" * 60)

    quantity_dict = {}
    remaining_parts = set(valid_parts)

    while remaining_parts:
        # Показываем оставшиеся детали
        print(f"\n  Осталось деталей: {len(remaining_parts)}")
        print("  Список оставшихся деталей:")
        for i, part_no in enumerate(sorted(remaining_parts), 1):
            print(f"    {i}. {part_no}")

        print("\n  Введите номер детали (или Enter чтобы установить 0 для всех оставшихся):")
        try:
            part_input = input("  → ").strip()
        except KeyboardInterrupt:
            print("\n\nПрограмма прервана пользователем (Ctrl+C)")
            sys.exit(0)

        # Если Enter - устанавливаем 0 для всех оставшихся
        if part_input == '':
            for part_no in remaining_parts:
                quantity_dict[part_no] = 0
                print(f"    • {part_no} → количество: 0 (установлено автоматически)")
            break

        # Проверяем, существует ли такой номер детали
        if part_input not in remaining_parts:
            print(f"  ОШИБКА: Деталь с номером '{part_input}' не найдена в списке!")
            print("  Пожалуйста, введите номер из списка выше.")
            continue

        # Запрашиваем количество
        while True:
            try:
                qty_input = input(f"  Введите количество для детали {part_input}: ").strip()
            except KeyboardInterrupt:
                print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                sys.exit(0)

            if qty_input == '':
                quantity = 0
                print("    → Установлено количество: 0")
                break
            else:
                try:
                    quantity = int(qty_input)
                    if quantity < 0:
                        print("  Количество не может быть отрицательным. Попробуйте снова.")
                        continue
                    print(f"    → Установлено количество: {quantity}")
                    break
                except ValueError:
                    print("  Ошибка: Введите целое число или нажмите Enter.")
                    continue

        quantity_dict[part_input] = quantity
        remaining_parts.remove(part_input)

    # Выводим итоговую таблицу
    print("\n" + "=" * 60)
    print("  ИТОГОВЫЕ ЗНАЧЕНИЯ КОЛИЧЕСТВА ДЕТАЛЕЙ В SS:")
    print("=" * 60)
    for part_no in valid_parts:
        qty = quantity_dict.get(part_no, 0)
        print(f"    {part_no}: {qty} шт.")
    print("=" * 60)

    print(f"\n  Ввод количества завершён. Обработано деталей: {len(quantity_dict)} из {total_parts}")
    return quantity_dict


def get_supplier_localization_status(df_bp_new, bp_number):
    """
    Интерактивный ввод статуса локализации для каждого уникального поставщика.

    Доступные статусы:
    - 1: Да (локальный поставщик) → "Да\nYes"
    - 2: Нет (зарубежный поставщик) → "Нет\nNo"
    - Enter: пустое значение

    Аргументы:
        df_bp_new (pd.DataFrame): DataFrame с данными BP.
        bp_number (str): Номер BP для вывода.

    Возвращается:
        pd.DataFrame: DataFrame с добавленной колонкой 'Localization'.
    """
    print(f"\n  Ввод статуса локализации поставщиков для BP {bp_number}")

    if 'Supplier Name (RUS)' not in df_bp_new.columns:
        print("  Колонка 'Supplier Name (RUS)' отсутствует. Создаём пустую колонку для статуса локализации.")
        df_bp_new['Localization'] = ''
        return df_bp_new

    # Получаем уникальных поставщиков
    unique_suppliers = df_bp_new['Supplier Name (RUS)'].drop_duplicates()
    unique_suppliers = [s for s in unique_suppliers if s != '-' and s != '' and s != 'nan']

    if len(unique_suppliers) == 0:
        print("  Нет уникальных поставщиков для указания статуса локализации")
        df_bp_new['Localization'] = ''
        return df_bp_new

    print(f"\n  Найдено уникальных поставщиков: {len(unique_suppliers)}")
    print("\n  Список поставщиков для указания статуса локализации:")
    print("-" * 60)
    for i, supplier in enumerate(unique_suppliers, 1):
        print(f"  {i}. {supplier}")
    print("-" * 60)

    print("\n  Доступные статусы локализации:")
    print("    1 - Да (локальный поставщик)")
    print("    2 - Нет (зарубежный поставщик)")
    print("    (или Enter, чтобы оставить пустым)")

    # Создаем словарь для хранения статусов
    localization_statuses = {}

    print("\n  Введите статус локализации для каждого поставщика:")
    for i, supplier in enumerate(unique_suppliers, 1):
        print(f"\n  [{i}/{len(unique_suppliers)}] Поставщик: {supplier}")

        while True:
            try:
                choice = input("  Введите статус (1, 2 или Enter чтобы оставить пустым): ").strip()

                if choice == '':
                    status = ''
                    print("    → Статус не указан (будет пустым)")
                    break
                elif choice == '1':
                    status = 'Да\nYes'
                    print(f"    → Статус: {status}")
                    break
                elif choice == '2':
                    status = 'Нет\nNo'
                    print(f"    → Статус: {status}")
                    break
                else:
                    print("  Неверный выбор. Пожалуйста, введите число от 1, 2 или нажмите Enter.")
            except KeyboardInterrupt:
                print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                sys.exit(0)

        localization_statuses[supplier] = status

    # Добавляем колонку со статусом локализации
    df_bp_new['Localization'] = df_bp_new['Supplier Name (RUS)'].map(localization_statuses).fillna('')

    # Показываем результат
    print("\n  Результат добавления статусов локализации:")
    result_df = df_bp_new[['Supplier Name (RUS)', 'Localization']].drop_duplicates()
    for _, row in result_df.iterrows():
        status_display = row['Localization'] if row['Localization'] else ''
        print(f"    {row['Supplier Name (RUS)']} → {status_display}")

    return df_bp_new


def translate_production_part_disposal(df_bp_new, bp_number):
    """
    Интерактивный перевод значений колонки 'Production Part Disposal'
    Формат сохранения: "перевод\nоригинал"

    Аргументы:
        df_bp_new (pd.DataFrame): DataFrame с данными BP.
        bp_number (str): Номер BP для вывода.

    Возвращается:
        pd.DataFrame: DataFrame с обновлённой колонкой 'Production Part Disposal'.
    """
    print(f"\n  Перевод значений 'Production Part Disposal' для BP {bp_number}")

    if 'Production Part Disposal' not in df_bp_new.columns:
        print("  Колонка 'Production Part Disposal' отсутствует. Пропускаем шаг.")
        return df_bp_new

    # Получаем уникальные значения, исключая пустые и '-'
    unique_values = get_unique_non_empty_values(df_bp_new['Production Part Disposal'], 'Production Part Disposal')

    if len(unique_values) == 0:
        print("  Нет уникальных значений для перевода. Пропускаем шаг.")
        return df_bp_new

    print(f"\n  Найдено уникальных значений: {len(unique_values)}")
    print("\n  Список значений для перевода:")
    print("-" * 60)
    for i, value in enumerate(unique_values, 1):
        print(f"  {i}. {value}")
    print("-" * 60)

    print("\n  Инструкция:")
    print("  • Введите перевод требования")
    print("  • Нажмите Enter без ввода, чтобы оставить исходное значение без изменений")

    # Создаем словарь для хранения переводов
    translations = {}

    for i, value in enumerate(unique_values, 1):
        print(f"\n  [{i}/{len(unique_values)}] Исходное значение: {value}")

        while True:
            try:
                user_input = input("  Введите перевод (или Enter чтобы оставить без изменений): ").strip()

                if user_input == '':
                    translations[value] = value
                    print(f"    → Оставляем без изменений: {value}")
                else:
                    translations[value] = f"{user_input}\n{value}"
                    print(f"    → Сохранено: {user_input}\\n{value}")
                break
            except KeyboardInterrupt:
                print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                sys.exit(0)

    # Применяем переводы
    df_bp_new['Production Part Disposal'] = df_bp_new['Production Part Disposal'].map(translations).fillna(df_bp_new['Production Part Disposal'])

    # Показываем результат
    print("\n  Результат перевода 'Production Part Disposal':")
    result_df = df_bp_new[['Production Part Disposal']].drop_duplicates()
    for _, row in result_df.iterrows():
        display_value = row['Production Part Disposal'] if len(row['Production Part Disposal']) <= 60 else row['Production Part Disposal'][:57] + '...'
        print(f"    {display_value}")

    return df_bp_new


def translate_interchangeable(df_bp_new, bp_number):
    """
    Интерактивный перевод значений колонки 'Interchangeable'.
    Формат сохранения: "перевод\nоригинал"

    Аргументы:
        df_bp_new (pd.DataFrame): DataFrame с данными BP.
        bp_number (str): Номер BP для вывода.

    Возвращается:
        pd.DataFrame: DataFrame с обновлённой колонкой 'Interchangeable'.
    """
    print(f"\n  Перевод значений 'Interchangeable' для BP {bp_number}")

    if 'Interchangeable' not in df_bp_new.columns:
        print("  Колонка 'Interchangeable' отсутствует. Пропускаем шаг.")
        return df_bp_new

    # Получаем уникальные значения, исключая пустые и '-'
    unique_values = get_unique_non_empty_values(df_bp_new['Interchangeable'], 'Interchangeable')

    if len(unique_values) == 0:
        print("  Нет уникальных значений для перевода. Пропускаем шаг.")
        return df_bp_new

    print(f"\n  Найдено уникальных значений: {len(unique_values)}")
    print("\n  Список значений для перевода:")
    print("-" * 60)
    for i, value in enumerate(unique_values, 1):
        print(f"  {i}. {value}")
    print("-" * 60)

    print("\n  Инструкция:")
    print("  • Введите перевод требования")
    print("  • Нажмите Enter без ввода, чтобы оставить исходное значение без изменений")

    # Создаем словарь для хранения переводов
    translations = {}

    for i, value in enumerate(unique_values, 1):
        print(f"\n  [{i}/{len(unique_values)}] Исходное значение: {value}")

        while True:
            try:
                user_input = input("  Введите перевод (или Enter чтобы оставить без изменений): ").strip()

                if user_input == '':
                    translations[value] = value
                    print(f"    → Оставляем без изменений: {value}")
                else:
                    translations[value] = f"{user_input}\n{value}"
                    print(f"    → Сохранено: {user_input}\\n{value}")
                break
            except KeyboardInterrupt:
                print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                sys.exit(0)

    # Применяем переводы
    df_bp_new['Interchangeable'] = df_bp_new['Interchangeable'].map(translations).fillna(df_bp_new['Interchangeable'])

    # Показываем результат
    print("\n  Результат перевода 'Interchangeable':")
    result_df = df_bp_new[['Interchangeable']].drop_duplicates()
    for _, row in result_df.iterrows():
        display_value = row['Interchangeable'] if len(row['Interchangeable']) <= 60 else row['Interchangeable'][:57] + '...'
        print(f"    {display_value}")

    return df_bp_new


def process_bp_file(bp_filename, df_bom):
    """
    Обработка одного BP файла через 18 последовательных шагов.

    Каждый шаг включает:
    - Интерактивное взаимодействие с пользователем
    - Возможность отката (retry)
    - Сохранение состояния

    Аргументы:
        bp_filename (str): Имя файла BP для обработки.
        df_bom (pd.DataFrame): DataFrame с данными BOM для проверки наличия деталей.

    Возвращается:
        dict: Словарь вида {'bp_number': str, 'dataframe': pd.DataFrame}
              или None при критической ошибке.

    Шаги обработки:
        1. Загрузка BP файла
        2. Выбор нужных колонок
        3. Выбор статуса тех. изменения
        4. Заполнение пустых значений
        5. Ввод количества деталей в SS
        6. Перевод названий деталей
        7. Поиск официальных названий поставщиков
        8. Ввод статуса локализации поставщиков
        9. Фильтрация китайских символов
        10. Перевод описания к изменению
        11. Перевод решения к изменению
        12. Обработка цветов и Color Code
        13. Обработка рабочих центров
        14. Перевод требований по утилизации старых деталей
        15. Перевод требований по взаимозаменяемости
        16. Проверка наличия деталей в BOM
        17. Упорядочивание колонок
        18. Сохранение результата
    """
    print(f"\nОбработка файла: {bp_filename}")

    bp_number = bp_filename.replace('.xlsx', '')
    saved_state = None  # Инициализируем сохранённое состояние

    # Шаг 1: Загрузка BP файла
    print_step_header(1, 18, "Загрузка BP файла")
    df_bp = load_excel_file(bp_filename, "BP файл")
    if df_bp is None:
        return None
    show_dataframe_preview(df_bp, "Загрузка исходного BP файла")

    # Сохраняем состояние после Шага 1
    saved_state = save_state_before_step(df_bp)

    continue_flag, df_bp, saved_state = confirm_step("Загрузка BP файла", df_bp, saved_state)
    if not continue_flag:
        return process_bp_file(bp_filename, df_bom)

    # Шаг 2: Выбор нужных колонок
    print_step_header(2, 18, "Выбор нужных колонок")
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

    show_dataframe_preview(
        df_bp_new, "Выбор нужных колонок",
        focus_columns=['BP_No', 'Change', 'BOM Product', 'Update Type', 'Part No.', 'Part Name(CHN)']
    )

    # Шаг 3: Выбор статуса тех. изменения
    # Запрашиваем статус для BP файла (только один раз для каждого файла)
    # Проверяем, не задан ли уже статус (например, при retry)
    print_step_header(3, 18, "Выбор статуса тех. изменения")
    if 'Status' not in df_bp_new.columns or df_bp_new['Status'].iloc[0] == '-':
        status = get_bp_status(bp_number)
        df_bp_new['Status'] = status
    else:
        print(f"  Статус для BP {bp_number} уже задан: {df_bp_new['Status'].iloc[0]}")
    show_dataframe_preview(
        df_bp_new, "Выбор статуса тех. изменения",
        focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Update Type', 'Part No.', 'Part Name(CHN)']
    )

    continue_flag, df_bp_new, saved_state = confirm_step("Выбор нужных колонок", df_bp_new, saved_state)
    if not continue_flag:
        return process_bp_file(bp_filename, df_bom)

    # Шаг 4: Заполнение пустых значений
    print_step_header(4, 18, "Заполнение пустых значений")
    columns_with_replacements = fill_empty_values_with_dash(df_bp_new, available_cols)

    base_columns = ['BOM Product', 'Part No.', 'Part Name(CHN)']
    preview_columns = []
    for col in base_columns:
        if col in df_bp_new.columns and col not in preview_columns:
            preview_columns.append(col)
    for col in columns_with_replacements:
        if col in df_bp_new.columns and col not in preview_columns:
            preview_columns.append(col)

    if preview_columns:
        show_dataframe_preview(
            df_bp_new, "Заполнение пустых значений",
            focus_columns=preview_columns
        )
    else:
        show_dataframe_preview(
            df_bp_new, "Заполнение пустых значений",
            focus_columns=['BP_No', 'Status', 'Change','BOM Product', 'Part No.', 'Part Name(CHN)']
        )

    continue_flag, df_bp_new, saved_state = confirm_step(
        "Заполнение пустых значений", df_bp_new, saved_state
    )
    if not continue_flag:
        return process_bp_file(bp_filename, df_bom)

    # Шаг 5: Ввод количества деталей в SS
    print_step_header(5, 18, "Ввод количества деталей в SS")
    quantity_dict = get_quantity_in_ss(df_bp_new, bp_number)
    df_bp_new['Quantity in SS'] = df_bp_new['Part No.'].map(quantity_dict).fillna(0).astype(int)

    show_dataframe_preview(
        df_bp_new, "Ввод Quantity in SS",
        focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (CHN)', 'Quantity in SS']
    )

    continue_flag, df_bp_new, saved_state = confirm_step("Ввод Quantity in SS", df_bp_new, saved_state)
    if not continue_flag:
        return process_bp_file(bp_filename, df_bom)

    # Шаг 6: Перевод названий деталей
    while True:
        print_step_header(6, 18, "Перевод названий деталей")
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
            show_dataframe_preview(
                df_bp_new, "Перевод названий деталей (после)",
                focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)']
            )

        continue_flag, df_bp_new, saved_state = confirm_step(
            "Перевод названий деталей", df_bp_new, saved_state
        )
        if continue_flag:
            break
        # при retry продолжаем цикл с восстановленным состоянием

    # Шаг 7: Поиск официальных названий поставщиков
    while True:
        print_step_header(7, 18, "Поиск официальных названий поставщиков")
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
            show_dataframe_preview(
                df_bp_new, "Поиск официальных названий поставщиков (после)",
                focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Supplier Name (RUS)']
            )

        continue_flag, df_bp_new, saved_state = confirm_step("Поиск официальных названий поставщиков", df_bp_new, saved_state)
        if continue_flag:
            break

    # Шаг 8: Ввод статуса локализации поставщиков
    while True:
        print_step_header(8, 18, "Ввод статуса локализации поставщиков")
        df_bp_new = get_supplier_localization_status(df_bp_new, bp_number)
        show_dataframe_preview(
            df_bp_new, "Ввод статуса локализации поставщиков",
            focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Supplier Name (RUS)', 'Localization']
        )

        continue_flag, df_bp_new, saved_state = confirm_step(
            "Ввод статуса локализации поставщиков", df_bp_new, saved_state
        )
        if continue_flag:
            break

    # Шаг 9: Фильтрация китайских символов
    print_step_header(9, 18, "Фильтрация китайских символов")
    if 'Change Description' in df_bp_new.columns:
        df_bp_new['Change Description'] = df_bp_new['Change Description'].apply(filter_chinese_lines)
        print("  Колонка 'Change Description': фильтрация выполнена")
    if 'Solution' in df_bp_new.columns:
        df_bp_new['Solution'] = df_bp_new['Solution'].apply(filter_chinese_lines)
        print("  Колонка 'Solution': фильтрация выполнена")
    show_dataframe_preview(
        df_bp_new, "Фильтрация китайских символов",
        focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Change Description', 'Solution']
    )

    continue_flag, df_bp_new, saved_state = confirm_step("Фильтрация китайских символов", df_bp_new, saved_state)
    if not continue_flag:
        return process_bp_file(bp_filename, df_bom)

    # Шаг 10: Перевод описания
    while True:
        print_step_header(10, 18, "Перевод описания к изменению")
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
            show_dataframe_preview(
                df_bp_new, "Перевод описания (после)",
                focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Change Description (RUS)']
            )

        continue_flag, df_bp_new, saved_state = confirm_step("Перевод описания изменений", df_bp_new, saved_state)
        if continue_flag:
            break

    # Шаг 11: Перевод решения
    while True:
        print_step_header(11, 18, "Перевод решения к изменению")
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
            show_dataframe_preview(
                df_bp_new, "Перевод решения (после)",
                focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Solution (RUS)']
            )

        continue_flag, df_bp_new, saved_state = confirm_step("Перевод решения", df_bp_new, saved_state)
        if continue_flag:
            break

    # Шаг 12: Обработка цветов и Color Code
    while True:
        print_step_header(12, 18, "Обработка цветов и Color Code")

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

        show_dataframe_preview(
            df_bp_new, "Обработка цветов",
            focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Color Code', 'Color Name (RUS)']
        )

        continue_flag, df_bp_new, saved_state = confirm_step("Обработка цветов и Color Code", df_bp_new, saved_state)
        if continue_flag:
            break

    # Шаг 13: Обработка рабочих центров
    while True:
        print_step_header(13, 18, "Обработка рабочих центров")
        if 'Workcenter Name' in df_bp_new.columns:
            df_bp_new['Workcenter Name'] = df_bp_new['Workcenter Name'].apply(extract_parentheses_content)

            unique_wc = get_unique_non_empty_values(df_bp_new['Workcenter Name'], 'Workcenter Name')

            if len(unique_wc) > 0:
                wc_translations = interactive_translation(unique_wc, "рабочих центров")
                df_bp_new['Workcenter Name'] = df_bp_new['Workcenter Name'].map(wc_translations).fillna('-')
            else:
                print("  Все значения рабочих центров равны '-' или пустые. Перевод не требуется.")
            show_dataframe_preview(
                df_bp_new, "Обработка рабочих центров",
                focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Workcenter Name']
            )

        continue_flag, df_bp_new, saved_state = confirm_step("Обработка рабочих центров", df_bp_new, saved_state)
        if continue_flag:
            break

    # Шаг 14: Перевод требований по дальнешейму использованию или утилизации старых деталей
    while True:
        print_step_header(14, 18, "Перевод требований по утилизации старых деталей")
        df_bp_new = translate_production_part_disposal(df_bp_new, bp_number)
        show_dataframe_preview(
            df_bp_new, "Перевод требований по утилизации старых деталей",
            focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Production Part Disposal']
        )

        continue_flag, df_bp_new, saved_state = confirm_step(
            "Перевод по утилизации старых деталей", df_bp_new, saved_state
        )
        if continue_flag:
            break

    # Шаг 15: Перевод требований по взаимозаменяемости
    while True:
        print_step_header(15, 18, "Перевод требований по взаимозаменяемости")
        df_bp_new = translate_interchangeable(df_bp_new, bp_number)
        show_dataframe_preview(
            df_bp_new, "Перевод требований по взаимозаменяемости",
            focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Interchangeable']
        )

        continue_flag, df_bp_new, saved_state = confirm_step(
            "Перевод требований по взаимозаменяемости", df_bp_new, saved_state
        )
        if continue_flag:
            break

    # Шаг 16: Проверка наличия в BOM
    print_step_header(16, 18, "Проверка наличия деталей в BOM")
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

    show_dataframe_preview(
        df_bp_new, "Проверка наличия в BOM",
        focus_columns=['BP_No', 'Status', 'Change', 'BOM Product', 'Part No.', 'Part Name (RUS)', 'Is in BOM']
    )

    continue_flag, df_bp_new, saved_state = confirm_step("Проверка наличия в BOM", df_bp_new, saved_state)
    if not continue_flag:
        return process_bp_file(bp_filename, df_bom)

    # Шаг 17: Упорядочивание колонок
    print_step_header(17, 18, "Упорядочивание колонок")
    bp_columns_order = [
        'BP_No', 'Status', 'In Stock', 'New Part Available Date', 'BOM Product',
        'Change', 'Update Type', 'Is in BOM', 'Part No.', 'Part Name (RUS)', 'Quantity',
        'Quantity in SS', 'Workcenter No.', 'Workcenter Name', 'Production Part Disposal',
        'Interchangeable', 'Supplier Name (RUS)', 'Localization', 'Change Description (RUS)', 'Solution (RUS)',
        'Color Code', 'Color Name (RUS)'
    ]

    existing_cols = [col for col in bp_columns_order if col in df_bp_new.columns]
    missing_cols_in_order = set(bp_columns_order) - set(existing_cols)

    print(f"  Выбрано колонок для финального вывода: {len(existing_cols)} из {len(bp_columns_order)}")
    if missing_cols_in_order:
        print(f"  Отсутствуют в данных: {missing_cols_in_order}")

    df_bp_new = df_bp_new[existing_cols]
    show_dataframe_preview(
        df_bp_new, "Упорядочивание колонок (финальный результат)",
        focus_columns=['BP_No', 'Status', 'Change', 'Part No.', 'Part Name (RUS)', 'Is in BOM']
    )

    continue_flag, df_bp_new, saved_state = confirm_step("Упорядочивание колонок", df_bp_new, saved_state)
    if not continue_flag:
        return process_bp_file(bp_filename, df_bom)

    # Шаг 18: Сохранение результата
    print_step_header(18, 18, "Сохранение результата")

    bp_dataframe = {
        'bp_number': bp_number,
        'dataframe': df_bp_new,
    }

    return bp_dataframe


def main():
    """
    Главная функция модуля bp_refactoring.

    Выполняет пошаговую обработку Excel файлов BP.

    Этапы:
        1. Загрузка BOM файла (обязательно)
        2. Проверка новых BP в системе
        3. Поиск BP файлов
        4. Пошаговая обработка каждого найденного BP файла
        5. Возврат словаря обработанных DataFrame'ов

    При отсутствии новых BP предоставляет возможность ручного ввода
    номеров BP для обработки.

    Возвращается:
        dict: Словарь обработанных DataFrame'ов, где ключ - номер BP,
              значение - обработанный DataFrame, или None при отсутствии обработки.

    Исключения:
        SystemExit: при критических ошибках (отсутствие BOM файла и т.д.)
    """
    print("""
        ╔══════════════════════════════════════════════════════════════╗
        ║                                                              ║
        ║              ПОШАГОВАЯ ОБРАБОТКА EXCEL ФАЙЛОВ BP             ║
        ║                                                              ║
        ╚══════════════════════════════════════════════════════════════╝
        """)

    print("\nБудут обработаны следующие шаги для каждого BP файла:")
    print("   1. Статус тех. изменения (Breakpoint)")
    print("   2. Количество старых деталей на складе Safety Stock")
    print("   3. Перевод названий деталей")
    print("   4. Поиск официальных названий поставщиков")
    print("   5. Статусы локализации поставщиков")
    print("   6. Перевод описаний и решений")
    print("   7. Перевод цвета деталей")
    print("   8. Обработка рабочих центров")
    print("   9. Перевод требований по дальнейшему использованию или утилизации")
    print("   10. Перевод требований по взаимозаменяемости")
    print("   11. Проверка наличия деталей в BOM")

    wait_for_user()

    # ЭТАП 1: Загрузка BOM файла
    print("\n" + "=" * 60)
    print("ЭТАП 1: Загрузка BOM файла")
    print("=" * 60)
    df_bom = load_excel_file('bom.xlsx', "BOM файл")
    if df_bom is None:
        print("\nКритическая ошибка: Не найден или не загружен файл bom.xlsx!")
        print("Убедитесь, что файл bom.xlsx находится в той же папке и имеет правильный формат.")
        wait_for_user()
        sys.exit(1)

    # Пауза после загрузки BOM файла
    wait_for_user()

    # ЭТАП 2: Проверка новых BP
    print("\n" + "=" * 60)
    print("ЭТАП 2: Проверка новых BP в системе")
    print("=" * 60)

    new_bp_set = new_bp_check()

    if new_bp_set is None:
        print("\n  Ошибка при проверке новых BP. Продолжение невозможно.")
        wait_for_user()
        sys.exit(1)

    if len(new_bp_set) > 0:
        print("\n" + "!" * 60)
        print("  ВНИМАНИЕ: Найдены новые BP, которые отсутствуют в breakpoint_data.xlsx!")
        print("!" * 60)
        print("\n  Действия пользователя:")
        print("    1. Скачайте из системы G-BOM Excel файлы для следующих BP:")
        for bp in sorted(new_bp_set):
            print(f"       - {bp}")
        print("    2. Поместите скачанные файлы в текущую папку")
        print("    3. Убедитесь, что файлы имеют формат: BP<номер>.xlsx")
        print("\n  После скачивания файлов программа продолжит работу.")

        wait_for_user("\n  Нажмите Enter, когда все файлы будут скачаны и помещены в текущую папку...")

        # Проверяем, появились ли файлы
        expected_files = [f"{bp}.xlsx" for bp in new_bp_set]
        missing_files = []

        print("\n  Проверка наличия скачанных файлов:")
        for expected_file in expected_files:
            if os.path.exists(expected_file):
                print(f"    {expected_file} - найден")
            else:
                print(f"    {expected_file} - НЕ НАЙДЕН")
                missing_files.append(expected_file)

        if missing_files:
            print("\n  Предупреждение: Не все файлы найдены!")
            print("  Отсутствуют:", missing_files)
            proceed = input("\n  Продолжить с имеющимися файлами? (да/нет): ").strip().lower()
            if proceed != 'да':
                print("  Программа завершена. Скачайте недостающие файлы и запустите снова.")
                sys.exit(0)
    else:
        print("\n  Новых BP для обработки не найдено.")
        print("-" * 60)
        print("  ВОЗМОЖНОСТЬ РУЧНОГО ВВОДА BP:")
        print("  Если вы хотите обработать конкретные BP, которые уже есть в папке,")
        print("  вы можете указать их номера вручную.")
        print("-" * 60)

        manual_input = input("\n  Хотите указать BP номера для обработки вручную? (да/нет): ").strip().lower()

        if manual_input == 'да':
            print("\n  ИНСТРУКЦИЯ ПО ВВОДУ BP НОМЕРОВ:")
            print("    1. Вводите номера BP в формате: BP26002813 (с префиксом 'BP')")
            print("    2. После ввода каждого номера нажмите Enter")
            print("    3. Для завершения ввода оставьте строку пустой и нажмите Enter")
            print("    4. Пример: BP12345")
            print("-" * 60)

            manual_bp_files = []

            while True:
                try:
                    bp_input = input("\n  Введите номер BP (или Enter для завершения): ").strip().upper()

                    if bp_input == '':
                        if len(manual_bp_files) == 0:
                            print("  Не введено ни одного BP. Программа завершает работу.")
                            sys.exit(0)
                        break

                    # Проверяем формат: должен начинаться с BP и содержать только цифры после этого
                    if bp_input.startswith('BP') and bp_input[2:].isdigit():
                        # Добавляем .xlsx если нужно
                        bp_filename = bp_input if bp_input.endswith('.xlsx') else f"{bp_input}.xlsx"
                        
                        # Проверяем, существует ли файл в текущей папке
                        if os.path.exists(bp_filename):
                            manual_bp_files.append(bp_filename)
                            print(f"    → {bp_filename} добавлен в список для обработки")
                        else:
                            print(f"    ОШИБКА: Файл '{bp_filename}' не найден в текущей папке!")
                            print("    Убедитесь, что файл скачан и находится в текущей директории.")
                            continue
                    else:
                        print(f"    ОШИБКА: '{bp_input}' не является корректным номером BP")
                        print("    Используйте формат: BP26002813")
                        continue

                except KeyboardInterrupt:
                    print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                    sys.exit(0)

            # Проверяем, что введены файлы
            if not manual_bp_files:
                print("\n  Не введено ни одного BP. Программа завершает работу.")
                sys.exit(0)

            # Устанавливаем список файлов для обработки
            bp_files = manual_bp_files
            print("\n" + "=" * 60)
            print("  РУЧНОЙ ВВОД BP ЗАВЕРШЁН")
            print("=" * 60)
            print(f"  Добавлено BP файлов для обработки: {len(bp_files)}")
            for i, f in enumerate(bp_files, 1):
                print(f"    {i}. {f}")

            wait_for_user("\n  Нажмите Enter, чтобы продолжить...")

            # Переходим к обработке файлов (пропускаем этап поиска BP файлов)
            # Обработка каждого BP файла
            processed_results = {}

            for i, bp_file in enumerate(bp_files, 1):
                print("\n" + "=" * 60)
                print(f"ОБРАБОТКА BP ФАЙЛА {i}/{len(bp_files)}: {bp_file}")
                print("=" * 60)

                print(f"\nТекущий файл: {bp_file}")

                wait_for_user("\nНажмите Enter для начала обработки этого файла...")

                result = process_bp_file(bp_file, df_bom)
                if result:
                    processed_results[result['bp_number']] = result['dataframe']

                if i < len(bp_files):
                    wait_for_user("\nФайл обработан. Нажмите Enter для перехода к следующему файлу...")

            # Итоги
            print("\n" + "=" * 60)
            print("ОБРАБОТКА ЗАВЕРШЕНА")
            print("=" * 60)
            print(f"\nОбработано файлов: {len(processed_results)}/{len(bp_files)}")

            if processed_results:
                print("\nОбработанные Breakpoint'ы:")
                for bp_number in processed_results:
                    print(f"   - {bp_number}")
                return processed_results

            print("\nНе обработано ни одного файла.")
            return None

        else:
            print("\n  Программа завершает работу.")
            wait_for_user()
            sys.exit(0)

    # ЭТАП 3: Поиск BP файлов
    print("\n" + "=" * 60)
    print("ЭТАП 3: Поиск BP файлов")
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

    # Пауза после поиска всех BP файлов
    wait_for_user()

    # Обработка каждого BP файла
    processed_results = {}

    for i, bp_file in enumerate(bp_files, 1):
        print("\n" + "=" * 60)
        print(f"ОБРАБОТКА BP ФАЙЛА {i}/{len(bp_files)}: {bp_file}")
        print("=" * 60)

        print(f"\nТекущий файл: {bp_file}")

        wait_for_user("\nНажмите Enter для начала обработки этого файла...")

        result = process_bp_file(bp_file, df_bom)
        if result:
            processed_results[result['bp_number']] = result['dataframe']

        if i < len(bp_files):
            wait_for_user("\nФайл обработан. Нажмите Enter для перехода к следующему файлу...")

    # Итоги
    print("\n" + "=" * 60)
    print("ОБРАБОТКА ЗАВЕРШЕНА")
    print("=" * 60)
    print(f"\nОбработано файлов: {len(processed_results)}/{len(bp_files)}")

    if processed_results:
        print("\nОбработанные Breakpoint'ы:")
        for bp_number in processed_results:
            print(f"   - {bp_number}")
        return processed_results

    print("\nНе обработано ни одного файла.")
    return None

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
"""
BP Refactoring Tool - Точка входа

Этот модуль является главной точкой входа в приложение Breakpoint Refactoring Tool.

Он координирует выполнение всех этапов обработки технических изменений:
    1. Проверка наличия новых BP для скачивания из системы G-BOM
    2. Пошаговая обработка Excel файлов BP через модуль bp_refactoring
    3. Формирование итоговой сводной таблицы через модуль bp_summary
    4. Сохранение результата в Excel файл с форматированием

Модуль обеспечивает:
    - Интерактивное взаимодействие с пользователем через консоль
    - Сохранение состояния обработки между запусками
    - Автоматическую загрузку последнего обработанного файла
    - Форматирование выходного Excel файла с заданными стилями

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
import traceback
import warnings
from datetime import datetime
from typing import Optional

import pandas as pd
import numpy as np
from pandas.errors import EmptyDataError, ParserError
from openpyxl.utils.exceptions import InvalidFileException

from bp_processing import main as processing_main
from bp_summary import main as summary_main

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Для Windows консоли
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


def clear_screen():
    """
    Очистка экрана консоли
    Использует системную команду 'cls' для Windows или 'clear' для Unix-подобных систем.
    """
    os.system('cls' if os.name == 'nt' else 'clear')


def wait_for_user(prompt="\nНажмите Enter для продолжения..."):
    """
    Ожидает нажатия клавиши Enter от пользователя.

    Аргументы:
        prompt (str): Текст приглашения к вводу. По умолчанию содержит инструкцию.

    Обрабатывается:
        - KeyboardInterrupt (Ctrl+C) - запрашивает подтверждение перед завершением
        - EOFError - завершает программу при обнаружении конца ввода
    """
    while True:  # Цикл для повторной попытки ввода
        try:
            user_input = input(prompt)
            return user_input  # Успешный ввод
        except KeyboardInterrupt:
            print()  # Переход на новую строку
            while True:
                confirm_word = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                if confirm_word == 'да':
                    print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                    sys.exit(0)
                elif confirm_word == 'нет':
                    print("\nПродолжаем работу...")
                    break  # Выходим из внутреннего цикла и продолжаем внешний
                else:
                    print("Пожалуйста, введите 'да' или 'нет'")
            # После break из внутреннего цикла, продолжаем внешний цикл
            # То есть снова показываем prompt и ждём ввод
            continue
        except EOFError:
            print("\n\nОбнаружен конец ввода. Программа завершена.")
            sys.exit(0)


def find_latest_breakpoint_file(file_prefix: str = 'breakpoint_data') -> Optional[str]:
    """
    Находит файл breakpoint_data с самой поздней датой в имени.
    Функция ищет в текущей директории файлы, соответствующие шаблону:
    ГГГГ-ММ-ДД_breakpoint_data.xlsx и возвращает самый свежий по дате.
    
    Аргументы:
        file_prefix (str): Префикс имени файла для поиска.
                           По умолчанию 'breakpoint_data'.
    
    Возвращается:
        Имя самого свежего файла или None, если файлы не найдены
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
    latest_file = matching_files[0][1]

    print(f"  Найден последний файл: {latest_file} (от {matching_files[0][0].strftime('%d.%m.%Y')})")
    if len(matching_files) > 1:
        print(f"  Всего найдено файлов с историей: {len(matching_files)}")

    return latest_file


def normalize_breakpoint_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Нормализует DataFrame из файла breakpoint_data:
    - Приводит колонки к правильным типам данных
    - Заполняет пустые значения символом '-' (ВКЛЮЧАЯ колонки с датами)

    Аргументы:
        df (pd.DataFrame): DataFrame для нормализации.

    Возвращается:
        pd.DataFrame: Нормализованный DataFrame.

    Примечание:
        Функция создаёт копию DataFrame, не изменяя оригинал.
        Все колонки приводятся к строковому типу для единообразия,
        чтобы пустые значения можно было заменить на '-'.
    """
    if df is None or df.empty:
        return df

    df_normalized = df.copy()
    columns_processed = []

    # Определяем числовые колонки (они будут преобразованы в числа, затем в строки)
    numeric_columns = [
        'Quantity', 'Quantity in SS', 'Quantity per Vehicle Before',
        'Quantity per Vehicle After', 'Quantity per Box Before',
        'Quantity per Box After', 'Quantity batches in SS',
        'Batches for old parts using out'
    ]

    # Колонки с датами (будут преобразованы в строки с заменой пустых на '-')
    datetime_columns = ['Change Date', 'New Part Available Date']

    # 1. Обработка числовых колонок
    for col in numeric_columns:
        if col in df_normalized.columns:
            # Преобразуем в числовой тип (ошибки -> NaN)
            df_normalized[col] = pd.to_numeric(df_normalized[col], errors='coerce')
            # Замена Inf на 0
            if np.isinf(df_normalized[col]).any():
                inf_count = np.isinf(df_normalized[col]).sum()
                df_normalized[col] = df_normalized[col].replace([np.inf, -np.inf], 0)
                print(f"  Колонка '{col}': {inf_count} значений Inf заменено на 0")
            # Заполняем NaN нулями
            nan_count = df_normalized[col].isna().sum()
            if nan_count > 0:
                df_normalized[col] = df_normalized[col].fillna(0)
                print(f"  Колонка '{col}': {nan_count} пустых значений заменено на 0")
            columns_processed.append(col)

    # 2. Обработка колонок с датами (преобразуем в строки с '-')
    for col in datetime_columns:
        if col in df_normalized.columns:
            # Сначала пробуем преобразовать в datetime для валидации
            # Затем приводим к строковому формату
            temp_series = pd.to_datetime(df_normalized[col], errors='coerce')

            # Форматируем даты в строку ГГГГ-ММ-ДД, а NaT заменяем на '-'
            df_normalized[col] = temp_series.apply(
                lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '-'
            )

            # Считаем количество замен
            empty_count = (df_normalized[col] == '-').sum()
            if empty_count > 0:
                print(f"  Колонка '{col}': {empty_count} пустых значений заменено на '-'")

            columns_processed.append(col)

    # 3. Обработка всех остальных колонок (строковые)
    for col in df_normalized.columns:
        if col not in columns_processed:
            # Приводим к строковому типу
            df_normalized[col] = df_normalized[col].astype(str)

            # Заменяем пустые значения и 'nan' на '-'
            empty_mask = (
                (df_normalized[col].str.strip() == '') |
                (df_normalized[col].str.strip() == 'nan') |
                (df_normalized[col].str.strip() == 'None') |
                (df_normalized[col].str.strip() == 'NaT')
            )
            empty_count = empty_mask.sum()

            if empty_count > 0:
                df_normalized.loc[empty_mask, col] = '-'
                print(f"  Колонка '{col}': {empty_count} пустых значений заменено на '-'")

            columns_processed.append(col)

    print(f"\n  Нормализация завершена. Обработано колонок: {len(columns_processed)}")
    return df_normalized


def load_breakpoint_data(file_prefix: str = 'breakpoint_data') -> Optional[pd.DataFrame]:
    """
    Загружает самый свежий файл breakpoint_data с датой в имени.

    Функция автоматически находит последний сохранённый файл с историей
    и загружает его как DataFrame. Заголовки ожидаются в 3-й строке файла.

    Аргументы:
        file_prefix (str): Префикс имени файла для поиска.
                           По умолчанию 'breakpoint_data'.

    Возвращается:
        Optional[pd.DataFrame]: DataFrame с данными из файла или None,
                                если файлы не найдены или произошла ошибка загрузки.

    Примечание:
        При ошибке загрузки выводится сообщение и возвращается None,
        что сигнализирует о необходимости создания нового файла.
    """

    latest_file = find_latest_breakpoint_file(file_prefix)

    if latest_file is None:
        print("  Внимание: Не найдено ни одного файла breakpoint_data с историей")
        print("  Будет создан новый файл")
        return None

    try:
        # Загружаем с указанием строки заголовка (3-я строка = header=2)
        df = pd.read_excel(latest_file, sheet_name='pivot', header=2)
        print(f"  Файл '{latest_file}' загружен успешно")
        print(f"  В нём уже {len(df)} строк")

        # приводим данные к требуемым типам
        print("\n  Выполняется нормализация данных...")
        df = normalize_breakpoint_data(df)

        return df

    except FileNotFoundError:
        print(f"  Ошибка: Файл '{latest_file}' не найден")
        print("  Будет создан новый файл")
        return None
    except PermissionError:
        print(f"  Ошибка: Нет прав для чтения файла '{latest_file}'")
        print("  Закройте файл, если он открыт в Excel, и попробуйте снова.")
        print("  Будет создан новый файл")
        return None
    except EmptyDataError:
        print(f"  Ошибка: Файл '{latest_file}' пуст")
        print("  Будет создан новый файл")
        return None
    except ParserError as e:
        print(f"  Ошибка: Файл '{latest_file}' повреждён или имеет неверный формат: {e}")
        print("  Будет создан новый файл")
        return None
    except InvalidFileException:
        print(f"  Ошибка: Файл '{latest_file}' не является корректным Excel файлом")
        print("  Будет создан новый файл")
        return None
    except ValueError as e:
        if "Excel file format cannot be determined" in str(e):
            print(f"  Ошибка: Не удалось определить формат файла '{latest_file}'")
            print("  Убедитесь, что файл имеет расширение .xlsx или .xls")
        else:
            print(f"  Ошибка при загрузке файла '{latest_file}': {e}")
        print("  Будет создан новый файл")
        return None
    except Exception as e:
        # Непредвиденная ошибка
        print(f"  НЕПРЕДВИДЕННАЯ ОШИБКА при загрузке файла '{latest_file}': {e}")
        print(f"  Тип ошибки: {type(e).__name__}")
        print("  Будет создан новый файл")
        return None


def save_excel_with_formatting(
        df_new_data: pd.DataFrame,
        output_filename: str,
        sheet_name: str = 'pivot'
    ) -> bool:
    """
    Сохраняет DataFrame в Excel с форматированием (цвета, ширина колонок) с помощью xlsxwriter

    Функция применяет сложное форматирование к выходному Excel файлу:
        - Объединение ячеек для заголовков "ДЛЯ КЛАДОВЩИКОВ"
        - Цветовая схема для различных групп колонок (HEX 0F243E, FDE9D9, белый)
        - Настройка ширины колонок согласно спецификации
        - Перенос текста для колонок с длинным содержимым

    Требования к форматированию:
        - Шрифт: Arial, размер 10, чёрный (для всех ячеек)
        - 1-я строка: объединение колонок E-P с текстом "ДЛЯ КЛАДОВЩИКОВ", выравнивание по центру
        - 1-я строка: объединение колонок AC-AD с текстом "ДЛЯ КЛАДОВЩИКОВ", выравнивание по центру
        - 2-я и 3-я строка (заголовки), колонки с A по AM: шрифт белый, полужирный, Arial 10, заливка HEX 0F243E
        - Колонки с E по P: заливка HEX FDE9D9
        - Колонки AC и AD: заливка HEX FDE9D9
        - Остальные колонки: белая заливка
        - Колонки AI-AJ, AM: обязательный перенос по словам

    Ширина колонок:
        - A-H, J-O, AF, AH, AK-AL: 25
        - I, P, AE, AG: 80
        - Q, S, U-AB: 30
        - R, T, AC-AD: 50
        - AI-AJ, AM: 100 (с переносом слов)

    Аргументы:
        df_new_data (pd.DataFrame): DataFrame с данными для сохранения
        output_filename (str): Имя выходного файла
        sheet_name (str): Имя листа в Excel. По умолчанию 'pivot'

    Возвращается:
        bool: True если сохранение успешно, False при ошибке

    Примечания:
        - Требуется установленный пакет xlsxwriter
        - При отсутствии xlsxwriter выполняется сохранение без форматирования
        - При PermissionError или других ошибках выполняется fallback сохранение
    """
    try:
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            df_new_data.to_excel(writer, sheet_name=sheet_name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            # Формат для объединённых ячеек (1-я строка) - белый шрифт, полужирный, заливка 0F243E
            merged_cell_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'bold': True,
                'font_color': 'white',
                'bg_color': '#0F243E',
                'valign': 'vcenter',
                'align': 'center',
                'text_wrap': True,
                'border': 1
            })

            # Основной формат для данных (Arial, 10, чёрный, белый фон)
            data_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'font_color': 'black',
                'bg_color': '#FFFFFF',
                'valign': 'top',
            })

            # Формат для заголовков (2-я и 3-я строка) - белый шрифт, полужирный, заливка 0F243E
            header_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'bold': True,
                'font_color': 'white',
                'bg_color': '#0F243E',
                'valign': 'center',
                'align': 'center',
                'text_wrap': True,
                'border': 1
            })

            # Формат для колонок E-P (индексы 4-15) - заливка FDE9D9
            columns_e_p_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'font_color': 'black',
                'bg_color': '#FDE9D9',
                'valign': 'top',
            })

            # Формат для колонок AC-AD (индексы 28-29) - заливка FDE9D9
            columns_ac_ad_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'font_color': 'black',
                'bg_color': '#FDE9D9',
                'valign': 'top',
            })

            # Формат для длинного текста с переносом (обычные колонки)
            wrap_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'font_color': 'black',
                'bg_color': '#FFFFFF',
                'text_wrap': True,
                'valign': 'top',
            })

            # Формат для длинного текста в колонках E-P
            wrap_e_p_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'font_color': 'black',
                'bg_color': '#FDE9D9',
                'text_wrap': True,
                'valign': 'top',
            })

            # Формат для длинного текста в колонках AC-AD
            wrap_ac_ad_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'font_color': 'black',
                'bg_color': '#FDE9D9',
                'text_wrap': True,
                'valign': 'top',
            })

            # Формат для колонок AI-AJ, AM - ОБЯЗАТЕЛЬНЫЙ ПЕРЕНОС, белая заливка
            columns_wrap_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'font_color': 'black',
                'bg_color': '#FFFFFF',
                'text_wrap': True,
                'valign': 'top',
            })

            # === КОНФИГУРАЦИЯ ШИРИНЫ КОЛОНОК ===
            column_widths = {
                0: 25, 1: 25, 2: 25, 3: 25, 4: 25, 5: 25, 6: 25, 7: 25,          # A-H: 25 (индексы 0-7)
                8: 80,                                                           # I: 80 (индекс 8)
                9: 25, 10: 25, 11: 25, 12: 25, 13: 25, 14: 25,                   # J-O: 25 (индексы 9-14)
                15: 80,                                                          # P: 80 (индекс 15)
                16: 30,                                                          # Q: 30 (индекс 16)
                17: 50,                                                          # R: 50 (индекс 17)
                18: 30,                                                          # S: 30 (индекс 18)
                19: 50,                                                          # T: 50 (индекс 19)
                20: 30, 21: 30, 22: 30, 23: 30, 24: 30, 25: 30, 26: 30, 27: 30,  # U-AB: 30 (индексы 20-27)
                28: 50, 29: 50,                                                  # AC-AD: 50 (индексы 28-29)
                30: 80,                                                          # AE: 80 (индекс 30)
                31: 25,                                                          # AF: 25 (индекс 31)
                32: 80,                                                          # AG: 80 (индекс 32)
                33: 25,                                                          # AH: 25 (индекс 33)
                34: 100, 35: 100,                                                # AI-AJ: 100 (индексы 34-35) - с переносом слов
                36: 25, 37: 25,                                                  # AK-AL: 25 (индексы 36-37)
                38: 100,                                                         # AM: 100 (индекс 38) - с переносом слов
            }

            # Устанавливаем ширину колонок
            for col_num, width in column_widths.items():
                if col_num < len(df_new_data.columns):
                    worksheet.set_column(col_num, col_num, width)

            # Определяем индексы колонок со специальной заливкой
            columns_e_p_indices = list(range(4, 16))      # колонки E-P (индексы 4-15)
            columns_ac_ad_indices = [28, 29]              # колонки AC, AD (индексы 28-29)
            columns_wrap_indices = [34, 35, 38]           # колонки AI, AJ, AM (индексы 34, 35, 38) - обязательный перенос

            # === 1-я строка: объединение ячеек ===
            if len(df_new_data.columns) > 15:
                worksheet.merge_range(0, 4, 0, 15, "ДЛЯ КЛАДОВЩИКОВ", merged_cell_format)

            if len(df_new_data.columns) > 29:
                worksheet.merge_range(0, 28, 0, 29, "ДЛЯ КЛАДОВЩИКОВ", merged_cell_format)

            # Заполняем остальные колонки 1-й строки
            for col_num in range(len(df_new_data.columns)):
                if col_num in range(4, 16) or col_num in [28, 29]:
                    continue
                worksheet.write(0, col_num, '', merged_cell_format)

            # === 2-я и 3-я строка: заголовки ===
            for col_num in range(min(len(df_new_data.columns), 39)):
                col_name = df_new_data.columns[col_num] if col_num < len(df_new_data.columns) else ''
                worksheet.write(1, col_num, col_name, header_format)

                if len(df_new_data) >= 1:
                    value = df_new_data.iloc[0, col_num] if col_num < len(df_new_data.columns) else ''
                    worksheet.write(2, col_num, value, header_format)

            # === Строки данных ===
            for row_num in range(len(df_new_data)):
                for col_num in range(len(df_new_data.columns)):
                    value = df_new_data.iloc[row_num, col_num]
                    is_long_text = isinstance(value, str) and len(value) > 50

                    if col_num in columns_wrap_indices:
                        # Колонки AI, AJ, AM - обязательный перенос, белая заливка
                        worksheet.write(row_num + 3, col_num, value, columns_wrap_format)
                    elif col_num in columns_e_p_indices:
                        # Колонки E-P - заливка FDE9D9
                        if is_long_text:
                            worksheet.write(row_num + 3, col_num, value, wrap_e_p_format)
                        else:
                            worksheet.write(row_num + 3, col_num, value, columns_e_p_format)
                    elif col_num in columns_ac_ad_indices:
                        # Колонки AC-AD - заливка FDE9D9
                        if is_long_text:
                            worksheet.write(row_num + 3, col_num, value, wrap_ac_ad_format)
                        else:
                            worksheet.write(row_num + 3, col_num, value, columns_ac_ad_format)
                    else:
                        # Остальные колонки - белая заливка
                        if is_long_text:
                            worksheet.write(row_num + 3, col_num, value, wrap_format)
                        else:
                            worksheet.write(row_num + 3, col_num, value, data_format)

        print(f"  Файл '{output_filename}' сохранён с форматированием")
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
        file_prefix: str = 'breakpoint_data'
    ) -> Optional[str]:
    """
    Сохраняет обработанный DataFrame с объединением с существующими данными.

    Функция выполняет:
        1. Загрузку самого свежего существующего файла (если есть)
        2. Объединение существующих и новых данных
        3. Сохранение объединённого DataFrame в файл с датой в имени

    Все файлы сохраняются с префиксом ГГГГ-ММ-ДД_ для обеспечения истории.
    При следующем запуске автоматически загружается самый свежий файл.

    Аргументы:
        df_new_data (pd.DataFrame): Новый DataFrame для сохранения.
        file_prefix (str): Префикс имени файла. По умолчанию 'breakpoint_data'.

    Возвращается:
        Optional[str]: Имя сохранённого файла или None при ошибке сохранения.

    Примечания:
        - При несовпадении структуры колонок запрашивается подтверждение у пользователя
        - Объединение выполняется через pd.concat с ignore_index=True
    """
    # Список русских переводов колонок (порядок соответствует column_order)
    column_translation = [
        'Номер переключения', 'Статус переключения', 'Партия по плану', 'Дата выхода новой детали',
        'Партия по факту', 'Дата переключения', 'Модель', 'Номер "старой" детали до переключения',
        'Название "старой" детали до переключения',
        'Количество "старых" деталей до переключения на Safety Stock',
        'Количество партий со "старыми" деталями до переключения',
        'Конфигурация для использования остатка "старых" деталей',
        'Партии для использования остатка "старых" деталей',
        'Привод (трансмиссия)', 'Номер "новой" детали после переключения',
        'Название "новой" детали после переключения',
        'Код производственной линии "старой" детали до переключения',
        'Название производственной линии "старой" детали до переключения',
        'Код производственной линии "новой" детали после переключения',
        'Название производственной линии "новой" детали после переключения',
        'Количество "старых" деталей до переключения на 1 авто',
        'Количество "новых" деталей после переключения на 1 авто',
        'Количество "старых" деталей до переключения на 1 ящик',
        'Количество "новых" деталей после переключения на 1 ящик',
        'Размеры ящика (Д-Ш-В) до переключения',
        'Размеры ящика (Д-Ш-В) после переключения',
        'Размеры поддона (Д-Ш-В) до переключения',
        'Размеры поддона (Д-Ш-В) после переключения',
        'Использование "старых" деталей до переключения в производстве',
        'Взаимозаменяемость "старых/новых" деталей до/после переключения',
        'Название поставщика до переключения', 'Локализация до переключения',
        'Название поставщика после переключения', 'Локализация после переключения',
        'Описание переключения', 'Решение переключения', 'Код цвета',
        'Название цвета', 'Комментарии',
    ]

    # Загружаем самый свежий существующий файл (если есть)
    df_existing = load_breakpoint_data(file_prefix)

    # Объединяем данные
    if df_existing is not None and not df_existing.empty:
        print(f"  Объединение: {len(df_existing)} существующих строк + {len(df_new_data)} новых строк")

        # Проверяем, что колонки совпадают
        if list(df_existing.columns) != list(df_new_data.columns):
            print("  ВНИМАНИЕ: Структура колонок не совпадает!")
            print(f"    Существующие колонки: {list(df_existing.columns)}")
            print(f"    Новые колонки: {list(df_new_data.columns)}")

            # Защищённый ввод для подтверждения объединения
            while True:
                try:
                    proceed = input("  Продолжить объединение? (да/нет): ").strip().lower()
                    if proceed == 'да':
                        break
                    elif proceed == 'нет':
                        break
                    else:
                        print("  Некорректный ввод. Пожалуйста, введите 'да' или 'нет'.")
                except KeyboardInterrupt:
                    print()
                    while True:
                        try:
                            confirm_word = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
                            if confirm_word == 'да':
                                print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                                sys.exit(0)
                            elif confirm_word == 'нет':
                                print("\nПродолжаем работу...")
                                break
                            else:
                                print("Пожалуйста, введите 'да' или 'нет'")
                        except KeyboardInterrupt:
                            print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                            sys.exit(0)
                    continue

            if proceed != 'да':
                print("  Объединение отменено. Новые данные будут сохранены в отдельный файл.")
                current_date = datetime.now().strftime('%Y-%m-%d')
                filename = f"{current_date}_{file_prefix}_new.xlsx"
                # Вставляем строку с русскими переводами после заголовков
                russian_row = pd.DataFrame([column_translation], columns=df_new_data.columns)
                df_new_data = pd.concat([df_new_data.iloc[:1], russian_row, df_new_data.iloc[1:]], ignore_index=True)
                success = save_excel_with_formatting(df_new_data, filename)
                return filename if success else None

        # Объединяем DataFrame
        df_combined = pd.concat([df_existing, df_new_data], ignore_index=True)
        print(f"  Итого строк после объединения: {len(df_combined)}")
    else:
        df_combined = df_new_data
        print(f"  Создаётся новый файл с {len(df_combined)} строками")

    # Вставляем строку с русскими переводами на позицию 2 (индекс 1)
    russian_row = pd.DataFrame([column_translation], columns=df_combined.columns)
    df_combined = pd.concat([df_combined.iloc[:1], russian_row, df_combined.iloc[1:]], ignore_index=True)
    print("  Добавлена строка с русскими переводами колонок")

    # Получаем текущую дату
    current_date = datetime.now().strftime('%Y-%m-%d')

    # Сохраняем файл с датой
    filename = f"{current_date}_{file_prefix}.xlsx"
    print(f"\n  Сохранение файла: {filename}")
    success = save_excel_with_formatting(df_combined, filename)

    if success:
        print(f"\n  Файл успешно сохранён: {filename}")
        return filename
    else:
        print("\n  Ошибка при сохранении файла!")
        return None


def main():
    """
    Главная функция приложения Breakpoint Refactoring Tool.

    Выполняет пошаговый процесс обработки технических изменений:

    Этапы выполнения:
        1. Отображение информации о программе и инструкций для пользователя
        2. Ожидание подтверждения пользователя для начала работы
        3. Запуск обработки BP файлов через bp_refactoring.main()
        4. Формирование итоговой сводной таблицы через bp_summary.main()
        5. Сохранение результата в Excel файл с форматированием

    Требования к окружению:
        - Наличие следующих файлов в рабочей папке:
            - bp_list_2025-2026.xlsx
            - bom.xlsx
            - configuration.xlsx
            - Упаковочный_лист_<партия>.xlsx
            - BP*.xlsx (файлы технических изменений)

    Returns:
        None - при успешном выполнении программа завершается с кодом 0,
               при ошибках - с кодом 1.

    Исключения:
        KeyboardInterrupt: обрабатывается корректно с завершением программы
        Exception: перехватывается, выводится traceback и код возврата 1
    """
    clear_screen()

    print("""
        ╔══════════════════════════════════════════════════════════════╗
        ║               BREAKPOINT REFACTORING TOOL V1.0               ║
        ║                          ----------                          ║
        ║      ПРИЛОЖЕНИЕ ДЛЯ ОБРАБОТКИ ТЕХНИЧЕСКИХ ИЗМЕНЕНИЙ V1.0     ║
        ╚══════════════════════════════════════════════════════════════╝
        """)

    print("\nИнструкция:")
    print("   1. Программа предназначена для обработки технических изменений - Breakpoint (BP)")
    print("   2. Программа разделена на 3 этапа:")
    print("      2.1. Пошаговая Обработка Excel файлов BP")
    print("      2.2. Пошаговое формирование итоговой таблицы")
    print("           • Разделение деталей по парам 'До / После изменения'")
    print("      2.3. Сохранение итоговой таблицы в Excel файл 'ГГГГ-ММ-ДД_breakpoint_data.xlsx'")
    print("   3. Перед обрабаткой программа проверит имеются ли BP для скачивания из системы G-BOM")
    print("   4. Программа будет обрабатывать Excel файлы по одному")
    print("   5. Четко следуйте указаниям программы на каждом шаге")
    print("   6. После каждого шага Вам предоставляется возможность проверить внесенные изменения:")
    print("      6.1. Если внесенные изменения корректны, нажмите Enter")
    print("      6.2. Если внесенные изменения некорректны, введите 'retry'")
    print("      6.3. 'retry' отменит внесенные изменения и Вы сможете исправить неточность")
    print("      6.4. После исправления нажмите Enter")
    print("      6.5. Нажмите Enter, чтобы пропустить шаг и оставить его без изменений")
    print("   7. Нажмите клавиши Ctrl+C, чтобы прервать работу программы")
    print("\nТРЕБОВАНИЯ:")
    print("   1. Пользователь должен иметь доступ к системе G-BOM")
    print("      • Если у Вас нет доступа к системе G-BOM, обратитесь в PLD/ED:")
    print("         → Бариков Владимир / Barikov Vladimir")
    print("         → Ермолаева Мая / Ermolaeva Maya")
    print("   2. Пользователь должен иметь доступ к системе SCM")
    print("      • Если у Вас нет доступа к системе SCM, обратитесь в PLD/WL:")
    print("         → Федин Антон / Fedin Anton")
    print("   3. Пользователь должен иметь доступ к мессенджеру DingTalk")
    print("      • Информация по Breakpoint рассылается в 2 чатах DingTalk:")
    print("         → Break Point (BP) - админ: Алексеева Елизавета / Alekseeva Elizaveta (MD/PM)")
    print("         → Breakpoint PLD Info - админ: Бариков Владимир / Barikov Vladimir (PLD/ED)")
    print(f"   4. Все Excel файлы должны находится в рабочей папке → {os.getcwd()}")
    print("      • Список необходимых файлов для корректной работы программы:")
    print("         → 'ГГГГ-ММ-ДД_breakpoint_data.xlsx'")
    print("         → 'bp_list_2025-2026.xlsx'")
    print("         → 'bom.xlsx'")
    print("         → 'configuration.xlsx'")
    print("         → 'BP<номер>.xlsx'")
    print("         → 'Упаковочный_лист_<партия>.xlsx'")
    print("\n\n***В случае некорректной работы программы обращаться к разработчику:")
    print("      • В мессенджере DingTalk:")
    print("         → Бариков Владимир / Barikov Vladimir (PLD/ED)")

    wait_for_user()

    print("=" * 70)
    print("ЗАПУСК BP REFACTORING TOOL")
    print("=" * 70)

    # Шаг 1: Обработка BP файлов
    print("\n[1] Обработка BP файлов...")
    processed_results = processing_main()

    if not isinstance(processed_results, dict) or len(processed_results) == 0:
        print("\nОшибка: Не обработано ни одного BP файла!")
        sys.exit(1)

    # Шаг 2: Формирование итоговой таблицы
    print("\n[2] Формирование итоговой таблицы...")
    summary_df = summary_main(processed_results)

    # Шаг 3: Сохранение результата
    print("\n[3] Сохранение результата...")
    saved_file = save_processed_dataframe(summary_df, 'breakpoint_data')

    if saved_file:
        print(f"\n  Готово! Файл: {saved_file}")
    else:
        print("\n  Ошибка при сохранении!")
        sys.exit(1)


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(__file__))

    try:
        main()
    except KeyboardInterrupt:
        print("\n\nПрограмма прервана пользователем (Ctrl+C)")
        sys.exit(0)
    except Exception as e:
        print(f"\n\n{'=' * 60}")
        print("НЕПРЕДВИДЕННАЯ ОШИБКА")
        print("=" * 60)
        print(f"Ошибка: {e}")
        print(f"Тип ошибки: {type(e).__name__}")
        print("\nПожалуйста, сообщите разработчику следующую информацию:")
        print("-" * 60)
        traceback.print_exc()
        print("=" * 60)
        sys.exit(1)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
"""
BP Refactoring Tool - Точка входа
"""
import sys
import os
import re
import traceback
from datetime import datetime
from typing import Optional

import pandas as pd

from bp_refactoring import main as refactoring_main
from bp_summary import main as summary_main


def find_latest_breakpoint_file(file_prefix: str = 'breakpoint_data') -> Optional[str]:
    """
    Находит файл breakpoint_data с самой поздней датой в имени
    
    Args:
        file_prefix: префикс имени файла
    
    Returns:
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


def load_breakpoint_data(file_prefix: str = 'breakpoint_data') -> Optional[pd.DataFrame]:
    """
    Загружает самый свежий файл breakpoint_data с датой в имени
    Возвращает DataFrame или None, если файлы не найдены
    """
    latest_file = find_latest_breakpoint_file(file_prefix)

    if latest_file is None:
        print("  Внимание: Не найдено ни одного файла breakpoint_data с историей")
        print("  Будет создан новый файл")
        return None

    try:
        # Загружаем с указанием строки заголовка (3-я строка = header=2)
        df = pd.read_excel(latest_file, header=2)
        print(f"  Файл '{latest_file}' загружен успешно")
        print(f"  В нём уже {len(df)} строк")
        return df
    except Exception as e:
        print(f"  Ошибка при загрузке '{latest_file}': {e}")
        print("  Будет создан новый файл")
        return None


def save_excel_with_formatting(
        df_new_data: pd.DataFrame,
        output_filename: str,
        sheet_name: str = 'pivot'
    ) -> bool:
    """
    Сохраняет DataFrame в Excel с форматированием (цвета, ширина колонок) с помощью xlsxwriter

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

    Args:
        df_new_data: DataFrame с новыми данными
        output_filename: имя файла для сохранения
        sheet_name: имя листа

    Returns:
        True если сохранение успешно, False если ошибка
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
    Сохраняет обработанный DataFrame с форматированием
    Все файлы сохраняются с датой в имени.
    При следующем запуске автоматически загружается самый свежий файл.
    
    Args:
        df_new_data: новый DataFrame для сохранения
        file_prefix: префикс имени файла
    
    Returns:
        Имя сохранённого файла или None при ошибке
    """
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

            proceed = input("  Продолжить объединение? (да/нет): ").strip().lower()
            if proceed != 'да':
                print("  Объединение отменено. Новые данные будут сохранены в отдельный файл.")
                current_date = datetime.now().strftime('%Y-%m-%d')
                filename = f"{current_date}_{file_prefix}_new.xlsx"
                success = save_excel_with_formatting(df_new_data, filename)
                return filename if success else None

        # Объединяем DataFrame
        df_combined = pd.concat([df_existing, df_new_data], ignore_index=True)
        print(f"  Итого строк после объединения: {len(df_combined)}")
    else:
        df_combined = df_new_data
        print(f"  Создаётся новый файл с {len(df_combined)} строками")

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
    """Главная функция"""
    print("=" * 70)
    print("ЗАПУСК BP REFACTORING TOOL")
    print("=" * 70)

    # Шаг 1: Обработка BP файлов
    print("\n[1] Обработка BP файлов...")
    processed_results = refactoring_main()

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
        print("\n\nПрограмма прервана пользователем")
        sys.exit(0)
    except Exception as e:
        print(f"\n\nОшибка: {e}")
        traceback.print_exc()
        sys.exit(1)

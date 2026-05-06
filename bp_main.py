#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BP Refactoring Tool - Точка входа
"""
import sys
import os
import traceback
from datetime import datetime
from typing import Optional, Tuple

import pandas as pd

from bp_refactoring import main as refactoring_main
from bp_summary import main as summary_main


def save_excel_file(
        df: pd.DataFrame,
        filename: str
    ) -> Tuple[bool, str]:
    """
    Сохранение Excel файла с проверкой на дубликаты
    Возвращает (успех, реальное_имя_файла)
    """
    if df is None or df.empty:
        print(f"Внимание: Нет данных для сохранения в {filename}")
        return False, filename

    actual_filename = filename

    if os.path.exists(actual_filename):
        print(f"Файл {actual_filename} уже существует!")
        name, ext = actual_filename.rsplit('.', 1)
        actual_filename = f"{name}_v2.{ext}"
        print(f"Сохраняем как {actual_filename}")

    try:
        df.to_excel(actual_filename, index=False)
        print(f"Файл сохранен: {actual_filename}")
        return True, actual_filename
    except PermissionError:
        print(f"Ошибка: Нет прав для записи в файл '{actual_filename}'")
        print("Закройте файл, если он открыт в Excel, и попробуйте снова.")
        return False, actual_filename
    except OSError as e:
        print(f"Ошибка при сохранении: {e}")
        return False, actual_filename
    except Exception as e:
        print(f"Непредвиденная ошибка при сохранении: {e}")
        return False, actual_filename


def save_processed_dataframe(
        df: pd.DataFrame,
        file_prefix: str = 'summary_breakpoint'
    ) -> Optional[str]:
    """
    Сохраняет обработанный DataFrame
    Возвращает имя сохранённого файла или None
    """
    current_date = datetime.now().strftime('%Y-%m-%d')
    filename = f"{current_date}_{file_prefix}.xlsx"

    success, actual_filename = save_excel_file(df, filename)
    return actual_filename if success else None


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

    # Шаг 3: Сохранение
    print("\n[3] Сохранение результата...")
    saved_file = save_processed_dataframe(summary_df, 'breakpoint_data')

    if saved_file:
        print(f"\nГотово! Файл: {saved_file}")
    else:
        print("\nОшибка при сохранении!")
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

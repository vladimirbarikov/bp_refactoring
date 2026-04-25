#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BP Refactoring Tool - Точка входа
"""
import sys
import os
from datetime import datetime
from typing import Optional

import pandas as pd

from bp_refactoring import main as refactoring_main
from summary_breakpoint_table import main as summary_main


def save_excel_file(df: pd.DataFrame, filename: str) -> bool:
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


def save_processed_dataframe(df: pd.DataFrame, file_prefix: str = 'summary_breakpoint') -> Optional[str]:
    """
    Сохраняет обработанный DataFrame
    Возвращает имя сохранённого файла или None
    """
    current_date = datetime.now().strftime('%Y-%m-%d')
    filename = f"{current_date}_{file_prefix}.xlsx"

    success = save_excel_file(df, filename)
    return filename if success else None


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
    saved_file = save_processed_dataframe(summary_df, 'summary_breakpoint')

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
        import traceback
        traceback.print_exc()
        sys.exit(1)

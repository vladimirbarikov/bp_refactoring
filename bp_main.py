#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BP Refactoring Tool - Точка входа
"""
import sys
import os
from datetime import datetime

from bp_refactoring import main as refactoing_main
from summary_breakpoint_table import main as summary_main


# def save_excel_file(df, filename):
#     """Сохранение Excel файла с проверкой на дубликаты"""
#     if df is None or df.empty:
#         print(f"Внимание: Нет данных для сохранения в {filename}")
#         return False

#     if os.path.exists(filename):
#         print(f"Файл {filename} уже существует!")
#         name, ext = filename.rsplit('.', 1)
#         new_filename = f"{name}_v2.{ext}"
#         print(f"Сохраняем как {new_filename}")
#         filename = new_filename

#     try:
#         df.to_excel(filename, index=False)
#         print(f"Файл сохранен: {filename}")
#         return True
#     except PermissionError:
#         print(f"Ошибка: Нет прав для записи в файл '{filename}'")
#         print("Закройте файл, если он открыт в Excel, и попробуйте снова.")
#         return False
#     except OSError as e:
#         print(f"Ошибка при сохранении: {e}")
#         return False
#     except Exception as e:
#         print(f"Непредвиденная ошибка при сохранении: {e}")
#         return False


# def save_processed_dataframe(df, bp_number):
#     """
#     Сохраняет обработанный DataFrame BP файла
#     Возвращает имя сохранённого файла или None
#     """
#     current_date = datetime.now().strftime('%Y-%m-%d')
#     filename = f"{current_date}_{bp_number}_refactored.xlsx"

#     # Используем универсальную функцию save_excel_file
#     success = save_excel_file(df, filename)
#     return filename if success else None


if __name__ == "__main__":
    # Добавляем текущую папку в путь поиска модулей
    sys.path.insert(0, os.path.dirname(__file__))

    # Запускаем main из bp_refactoring
    processed_results = refactoing_main()

    if isinstance(processed_results, dict):
        summary_main(processed_results)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
# pylint: disable=import-outside-toplevel
"""
BP Refactoring Tool - Точка входа (V2.0)

Этот модуль является главной точкой входа в приложение Breakpoint Refactoring Tool.

Он выполняет оркестрацию полного цикла обработки технических изменений (BP):

    ЭТАП 1: ПРОВЕРКА НОВЫХ НОМЕРОВ ТЕХНИЧЕСКИХ ИЗМЕНЕНИЙ (BREAKPOINT -> BP)
        - Аудит входящего отчёта 'breakpoint_report'
        - Аудит накопленной базы 'breakpoint_data'
        - Вычисление разности множеств (новые BP, требующие скачивания из G-BOM)
        - Ручной ввод дополнительных BP-номеров (опционально)

    ЭТАП 2: ПРОВЕРКА НАЛИЧИЯ BP ФАЙЛОВ (BP<номер>.xlsx)
        - Сканирование директории 'input_breakpoint_files'
        - Сверка ожидаемых файлов с физически присутствующими
        - Возможность докачать отсутствующие файлы или пропустить их

    ЭТАП 3: ПОТОКОВАЯ ОБРАБОТКА BP И СОХРАНЕНИЕ (YYYY-mm-dd_breakpoint_data.xlsx)
        - Пошаговая обработка каждого BP через py_lib.pipeline.processing
        - Автономный бэкап каждого BP в 'output_backup_files'
        - Инкрементальная транзакционная запись в глобальную базу 'breakpoint_data'

    ЭТАП 4: ИТОГИ
        - Сводка обработанных BP
        - Общее время работы программы

Модуль обеспечивает:
    - Интерактивное взаимодействие с пользователем через консоль
    - Сохранение состояния обработки между запусками
    - Безопасную инкрементальную запись результатов (транзакционность)
    - Автоматическое резервное копирование обработанных BP

Использование:
    python main.py

Версия: 2.0
Совместимость: Python 3.14.4+, Pandas 3.0.3+, OpenPyXL 3.1.5+
Поддержка: PLD Engineering Center
Дата создания: 2026-05-14
Дата изменения: 2026-09-24
Лицензия: MIT
Статус: Production
"""

import os
import sys
import time
import traceback
import warnings
from datetime import timedelta
from typing import Any, Dict, Optional

import pandas as pd

from py_lib.config.core import (
    # Пути к директориям
    # --- Входные данные ---
    INPUT_BREAKPOINT_DATA_DIR,
    INPUT_BREAKPOINT_REPORT_DIR,
    INPUT_BREAKPOINT_FILES_DIR,
    # --- Выходные данные ---
    OUTPUT_BACKUP_ROOT,
    OUTPUT_BREAKPOINT_DATA_ROOT,

    # Шаги обработки технических изменений
    BP_STEP_DESCRIPTIONS,

    # Префиксы файлов входных/выходных данных
    BP_FILE_PREFIX,
    BP_DATA_PREFIX,
    BP_REPORT_PREFIX,

    # Колонки с полными данными
    BP_REPORT_COLS,

    # Ключевые колонки
    BP_REPORT_KEY_COL,
    BP_DATA_KEY_COL,

    # Колонки по типам данных
    BP_DATA_INT_COLS,
    BP_DATA_DATETIME_COLS,
    BP_DATA_TRANSLATION_COLS,
)

from py_lib.engine.etl import (
    find_latest_excel_file,
    find_bp_files,
    read_excel_file,
    normalize_data,
    get_daily_report_path,
    save_backup,
    save_processed_dataframe,
)

from py_lib.pipeline.processing import process_bp_file

from py_lib.ui.interaction import (
    clear_screen,
    wait_for_user,
    ask_yes_no,
    ask_user_input,
)

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def is_valid_bp_number(
    bp_value: str
) -> bool:
    """
    Проверяет, что номер BP соответствует формату BP<YY><NNNNNN> и что
    годовая часть (первые 2 цифры числа) >= 26.

    Логика разбора:
        - Отбрасываем префикс 'BP' (без учёта регистра).
        - Оставшаяся часть должна быть числом длиной >= 2.
        - Первые две цифры числа — это 'YY' (год).
        - Валидно, если YY >= 26.

    Примеры:
        >>> is_valid_bp_number('BP27000822')  # YY='27' → True
        True
        >>> is_valid_bp_number('BP26000822')  # YY='26' → True
        True
        >>> is_valid_bp_number('BP25000822')  # YY='25' → False
        False
        >>> is_valid_bp_number('BP123')       # YY='12' → False
        False
        >>> is_valid_bp_number('BP')          # нет цифр → False
        False
        >>> is_valid_bp_number('-')           # мусор → False
        False

    Аргументы:
        bp_value (str): Номер BP из файла отчёта (например, 'BP26000822').

    Возвращается:
        bool: True, если номер валиден и его годовая часть >= 26.
    """
    if not isinstance(bp_value, str):
        return False

    value = bp_value.strip().upper()
    if not value.startswith(BP_FILE_PREFIX):
        return False

    digits = value[len(BP_FILE_PREFIX):]
    if not digits.isdigit() or len(digits) < 2:
        return False

    try:
        year_prefix = int(digits[:2])
    except (ValueError, TypeError):
        return False

    return year_prefix >= 26


def bp_report_check(
    bp_report_dir: str = INPUT_BREAKPOINT_REPORT_DIR
) -> set:
    """
    Поиск, чтение и аудит файла отчёта по техническим изменениям Breakpoint.

    Логика:
        1. Находим самый свежий файл отчёта в директории.
        2. Читаем файл, оставляя только колонку 'BP No.' (BP_REPORT_COLS).
        3. Извлекаем все значения номеров BP.
        4. Оставляем только валидные номера (годовая часть >= 26).
        5. Возвращаем множество номеров BP.

    Аргументы:
        bp_report_dir (str): Директория с файлами отчёта 'Breakpoint Report'.

    Возвращается:
        set: Множество валидных номеров BP (str).
    """
    print("\n  [Аудит] Проверка входящего отчёта 'Breakpoint Report'...")

    # 1. Поиск самого свежего файла отчёта
    latest_report_file = find_latest_excel_file(
        file_prefix=BP_REPORT_PREFIX,
        file_path=bp_report_dir
    )
    if not latest_report_file:
        print(f"  [Внимание] В папке '{bp_report_dir}' не найдено файлов 'Breakpoint Report'.")
        return set()

    # 2. Чтение файла (только нужные колонки)
    df_raw = read_excel_file(
        filename=latest_report_file,
        file_path=bp_report_dir,
        cols=BP_REPORT_COLS
    )
    if df_raw is None or df_raw.empty:
        return set()

    # 3. Нормализация: убираем мусор, NaN, None, пробелы
    df_clean = normalize_data(df_raw)

    # 4. Проверка наличия ключевой колонки
    col_bp = BP_REPORT_KEY_COL
    if col_bp not in df_clean.columns:
        print(f"  [Ошибка] В файле отчёта не найдена колонка '{col_bp}'.")
        return set()

    # 5. Извлекаем все значения номеров BP
    bp_series = df_clean[col_bp].astype(str)

    # 6. Фильтруем: оставляем только валидные номера (год >= 26)
    set_bp: set = set()
    invalid_count = 0

    for bp_value in bp_series:
        if is_valid_bp_number(bp_value):
            set_bp.add(bp_value.strip().upper())
        else:
            # Учитываем только непустые строки, отличные от '-'
            if bp_value and bp_value.strip() not in ('', '-'):
                invalid_count += 1

    print(f"  [Успех] Найдено {len(set_bp)} активных BP (год >= 26).")
    if invalid_count > 0:
        print(f"  [Информация] Отфильтровано невалидных значений: {invalid_count}.")

    return set_bp


def bp_data_check(
    history_dir: str = INPUT_BREAKPOINT_DATA_DIR
) -> set:
    """
    Поиск и чтение исторического файла базы данных breakpoint_data.xlsx.
    Нормализует данные перед детекцией ключевой колонки.
    """
    print("\n  [Аудит] Проверка накопленной базы данных breakpoint_data...")

    # 1. Поиск самого свежего файла истории
    latest_db_file = find_latest_excel_file(file_prefix=BP_DATA_PREFIX, file_path=history_dir)
    if not latest_db_file:
        print("  [История] Историческая база данных отсутствует. Все найденные BP будут считаться новыми.")
        return set()

    # 2. Чтение файла
    df_raw = read_excel_file(filename=latest_db_file, file_path=history_dir)
    if df_raw is None or df_raw.empty:
        return set()

    # 3.  Нормализация: убираем мусор, NaN, None, пробелы
    df_clean = normalize_data(df_raw)

    # 4. Проверка наличия ключевой колонки (без детекции — имя фиксировано)
    col_bp_no = BP_DATA_KEY_COL
    if col_bp_no not in df_clean.columns:
        print(f"  [Внимание] В файле истории не найдена колонка '{col_bp_no}'.")
        return set()

    # 5. Формирование множества существующих номеров
    bp_no_series = df_clean[col_bp_no]
    set_bp_no = {bp for bp in bp_no_series if bp and bp != '-'}

    # 6. Проверка на пустой результат (файл есть, но валидных записей нет)
    if not set_bp_no:
        print(f"  [Информация] Файл истории существует, но не содержит валидных номеров BP "
              f"в колонке '{col_bp_no}'. Все найденные BP будут считаться новыми.")
        return set()

    print(f"  [Внимание] В текущей базе данных уже содержатся записи по {len(set_bp_no)} уникальным BP.")
    return set_bp_no


def new_bp_check() -> Optional[set]:
    """
    Сверяет данные из входящего списка (breakpoint_report) и накопленной базы (breakpoint_data).
    Вычисляет разность множеств и возвращает сет с номерами исключительно новых BP.
    """
    # Вызываем наши новые изолированные модули аудита
    set_source_bp = bp_report_check()
    set_existing_bp = bp_data_check()

    # Находим разницу множеств: те, что есть в списке, но которых еще нет в базе
    new_bp_set = set_source_bp - set_existing_bp

    print("\n" + "=" * 60)
    print("  РЕЗУЛЬТАТ СИНХРОНИЗАЦИИ И ПРОВЕРКИ НОВЫХ BP:")
    print("=" * 60)
    print(f"  • Всего BP во входящем списке (активных): {len(set_source_bp)}")
    print(f"  • BP уже обработано и сохранено в базе:  {len(set_source_bp & set_existing_bp)}")
    print(f"  • НОВЫЕ BP (требуют скачивания и работы): {len(new_bp_set)}")

    if new_bp_set:
        print("\n  Список новых BP для скачивания из G-BOM:")
        for i, bp in enumerate(sorted(new_bp_set), 1):
            print(f"    {i}. {bp}")
    else:
        print("\n  [Отлично] Все технические изменения из списка уже обработаны. Новых файлов не требуется.")

    return new_bp_set


def inject_bp_to_global_report(
    df_bp_clean: pd.DataFrame,
    bp_number: str
) -> bool:
    """
    Выполняет безопасную инкрементальную дозапись одного обработанного DataFrame BP
    в глобальный исторический файл breakpoint_data.xlsx.

    Логика выбора источника истории (history_dir):
        1. Если в OUTPUT_BREAKPOINT_DATA_ROOT уже существует файл текущей сессии
           (YYYY-MM-DD_breakpoint_data.xlsx) — значит, предыдущие BP в этой сессии
           уже сохранялись. Читаем историю из OUTPUT, чтобы не потерять их при
           объединении.
        2. Иначе (первый BP сессии) — читаем историю из INPUT_BREAKPOINT_DATA_DIR
           (исходная база без новых BP).

    Это обеспечивает корректное накопление данных в рамках одной сессии:
    каждый следующий BP читает уже обновлённую базу из OUTPUT и добавляет к ней.

    Аргументы:
        df_bp_clean (pd.DataFrame): Обработанный DataFrame одного BP.
        bp_number (str):            Номер BP (для сообщений).

    Возвращается:
        bool: True при успешной записи, False при ошибке.
    """
    try:
        # Определяем, откуда читать историю для объединения.
        output_history_file = get_daily_report_path(
            file_prefix=BP_DATA_PREFIX,
            output_dir=OUTPUT_BREAKPOINT_DATA_ROOT,
        )

        if os.path.exists(output_history_file):
            # Предыдущий BP текущей сессии уже сохранён — читаем из OUTPUT,
            # чтобы не потерять его строки при следующем объединении.
            history_dir = OUTPUT_BREAKPOINT_DATA_ROOT
        else:
            # Первый BP сессии — читаем исходную базу из INPUT.
            history_dir = INPUT_BREAKPOINT_DATA_DIR

        export_path = save_processed_dataframe(
            df_new_data=df_bp_clean,
            file_prefix=BP_DATA_PREFIX,
            column_translation=BP_DATA_TRANSLATION_COLS,
            int_columns=BP_DATA_INT_COLS,
            datetime_columns=BP_DATA_DATETIME_COLS,
            history_dir=history_dir,
            output_dir=OUTPUT_BREAKPOINT_DATA_ROOT,
        )

        if export_path == "STRUCTURE_MISMATCH":
            print(f"  [Ошибка записи] Структура колонок файла {bp_number} нарушает формат базы данных!")
            return False
        elif export_path:
            return True

        return False
    except Exception as e:
        print(f"  [Системный сбой] Ошибка при попытке транзакционной записи {bp_number} на диск: {e}")
        return False


def main():
    """
    Главная управляющая функция модуля bp_refactoring.
    Оркестрирует аудит входящих файлов, валидацию и запуск конвейера обработки.

    Этапы оркестрации:
        ЭТАП 1: ПРОВЕРКА НОВЫХ НОМЕРОВ ТЕХНИЧЕСКИХ ИЗМЕНЕНИЙ (BP<номер>)
        ЭТАП 2: ПРОВЕРКА НАЛИЧИЯ BP ФАЙЛОВ (BP<номер>.xlsx)
        ЭТАП 3: ПОТОКОВАЯ ОБРАБОТКА BP И СОХРАНЕНИЕ (YYYY-mm-dd_breakpoint_data.xlsx)
        ЭТАП 4: ИТОГИ

    """
    clear_screen()

    # НАЧАЛО ЗАМЕРА ОБЩЕГО ВРЕМЕНИ ПРОГРАММЫ
    program_start_time = time.time()

    print("=" * 70)
    print("ЗАПУСК BP REFACTORING TOOL")
    print("=" * 70)

    print("""
            ╔══════════════════════════════════════════════════════════════╗
            ║               BREAKPOINT REFACTORING TOOL V2.0               ║
            ║                          ----------                          ║
            ║      ПРИЛОЖЕНИЕ ДЛЯ ОБРАБОТКИ ТЕХНИЧЕСКИХ ИЗМЕНЕНИЙ V2.0     ║
            ╚══════════════════════════════════════════════════════════════╝
            """)

    print("\nИнструкция:")
    print("   1. Программа предназначена для обработки технических изменений - Breakpoint (BP)")
    print("   2. Программа разделена на 4 этапа:")
    print("      ЭТАП 1: ПРОВЕРКА НОВЫХ НОМЕРОВ ТЕХНИЧЕСКИХ ИЗМЕНЕНИЙ (BP<номер>)")
    print("      ЭТАП 2: ПРОВЕРКА НАЛИЧИЯ BP ФАЙЛОВ (BP<номер>.xlsx)")
    print("      ЭТАП 3: ПОТОКОВАЯ ОБРАБОТКА BP И СОХРАНЕНИЕ (YYYY-mm-dd_breakpoint_data.xlsx)")
    print("      ЭТАП 4: ИТОГИ")
    print("   3. Четко следуйте указаниям программы на каждом этапе")
    print("   4. Вам предоставляется возможность проверить внесенные изменения:")
    print("      • Если внесенные изменения корректны, нажмите Enter")
    print("      • Если внесенные изменения некорректны, введите 'retry'")
    print("      • 'retry' отменит внесенные изменения и Вы сможете исправить неточность")
    print("      • После исправления нажмите Enter")
    print("      • Нажмите Enter, чтобы пропустить шаг и оставить его без изменений")
    print("   5. Программа будет обрабатывать BP файлы по одному")
    print("   6. Автосохранение:")
    print("      • Программа будет сохранять каждый обработанный BP в YYYY-mm-dd_breakpoint_data.xlsx")
    print("      • Программа будет сохранять бэкап в директорию 'output_backup_files' после обработки каждого BP")
    print("   7. Нажмите клавиши Ctrl+C, чтобы экстренно прервать работу программы")
    print("\nТРЕБОВАНИЯ:")
    print("   1. Пользователь должен иметь доступ к системе G-BOM")
    print("      • Если у Вас нет доступа к системе G-BOM, обратитесь в PLD/ED:")
    print("         → Ермолаева Майя / Ermolaeva Maya")
    print("   2. Пользователь должен иметь доступ к системе SCM")
    print("      • Если у Вас нет доступа к системе SCM, обратитесь в PLD/WL:")
    print("         → Федин Антон / Fedin Anton")
    print("   3. Пользователь должен иметь доступ к мессенджеру rLink")
    print("      • Информация по Breakpoint рассылается в 2 чатах rLink:")
    print("         → Break Points - админ: Бровкина Софья / Brovkina Sofya (MD/PM)")
    print("         → Breakpoint PLD Info - админ: Ермолаева Майя / Ermolaeva Maya (PLD/ED)")
    print(f"   4. Все Excel файлы должны находится в рабочей папке → {os.getcwd()}")
    print("      • Структура для корректной работы программы:")
    print("""
                .
                ├── BREAKPOINT_REFACTORING_TOOL_V2.0.exe
                ├── input_files
                │   ├── input_breakpoint_data_files
                │   │   └── YYYY-mm-dd_breakpoint_data.xlsx
                │   ├── input_breakpoint_files
                │   │   └── BP<номер>.xlsx
                │   ├── input_breakpoint_report_files
                │   │   └── YYYY-mm-dd_breakpoint_report.xlsx
                │   ├── input_configuration_files
                │   │   └── YYYY-mm-dd_configuration.xlsx
                │   └── input_packing_list_files
                │       └── Упаковочный лист партия <номер>.xlsx
                └──output_files
                    ├── output_backup_files
                    │   └──BP<номер>
                    │      └── BP<номер>.xlsx
                    └── output_breakpoint_data_files
                        └── YYYY-mm-dd_breakpoint_report.xlsx
        """)
    print("\n\n***В случае некорректной работы программы обращаться к разработчику:")
    print("      • В мессенджере rLink:")
    print("         → Бариков Владимир / Barikov Vladimir (PLD/ED)")

    wait_for_user()

    print("""
        ╔══════════════════════════════════════════════════════════════╗
        ║                                                              ║
        ║              ПОШАГОВАЯ ОБРАБОТКА EXCEL ФАЙЛОВ BP             ║
        ║                                                              ║
        ╚══════════════════════════════════════════════════════════════╝
        """)

    print("\nБудут обработаны следующие шаги для каждого BP файла:")

    for i, description in enumerate(BP_STEP_DESCRIPTIONS, 1):
        print(f"   ШАГ {i:>2}: {description}")

    wait_for_user()


    # =============================================================================
    # ЭТАП 1: ПРОВЕРКА НОВЫХ НОМЕРОВ ТЕХНИЧЕСКИХ ИЗМЕНЕНИЙ (BREAKPOINT -> BP)
    # =============================================================================
    print("\n" + "=" * 60)
    print("ЭТАП 1: ПРОВЕРКА НОВЫХ НОМЕРОВ ТЕХНИЧЕСКИХ ИЗМЕНЕНИЙ (BREAKPOINT -> BP)")
    print("=" * 60)

    # Запускаем аудит
    new_bp_set = new_bp_check()

    # Предохранитель на случай технического сбоя (если функция вернула None)
    if new_bp_set is None:
        print("\n  [Критическая ошибка] Сбой при синхронизации и проверке новых BP. Продолжение невозможно.")
        wait_for_user()
        sys.exit(1)

    # Формируем список файлов для автоматической обработки
    auto_bp_files = [f"{bp}.xlsx" for bp in new_bp_set]
    bp_target_dir = INPUT_BREAKPOINT_FILES_DIR

    # Развилка логики интерфейса
    if auto_bp_files:
        # Новые BP обнаружены
        print("\n" + "!" * 60)
        print("  [ВНИМАНИЕ]: Обнаружены новые BP, которые отсутствуют в вашей базе данных!")
        print("!" * 60)
        print("\n  Пожалуйста, выполните следующие действия:")
        print("    1. Скачайте из системы G-BOM необходимые Excel-файлы изменений.")
        print(f"    2. Поместите скачанные файлы в папку проекта: '{bp_target_dir}'")
        print("    3. Убедитесь, что файлы имеют строгое имя формата: BP<номер>.xlsx")
        print("\n  Программа приостановлена. Вы можете скопировать файлы прямо сейчас.")

        wait_for_user(f"\n  Поместите файлы в '{bp_target_dir}' и нажмите Enter для продолжения...")
    else:
        # Новых BP нет
        print("\n" + "─" * 60)
        print("  [Синхронизация] Новых технических изменений в системе G-BOM не обнаружено.")
        print("  Ваша локальная база данных находится в полностью актуальном состоянии!")
        print("─" * 60)
        wait_for_user()

    # Ручной ввод BP номеров
    print("\n" + "=" * 60)
    print("Дополнительный ручной ввод BP номеров (ОПЦИОНАЛЬНО)")
    print("=" * 60)

    if auto_bp_files:
        print(f"\n  Система автоматически запланировала к обработке {len(auto_bp_files)} BP.")
        prompt_msg = "  Хотите добавить дополнительные номера BP для обработки вручную?"
    else:
        print("\n  Поскольку новых автоматических изменений нет, конвейер пуст.")
        prompt_msg = "  Хотите указать номер конкретного файла BP для принудительной ручной обработки?"

    manual_bp_files = []

    # Запрашиваем у пользователя решение через универсальный UI-компонент
    if ask_yes_no(prompt_msg, default_yes=False):
        print("\n  ИНСТРУКЦИЯ ПО РУЧНОМУ ДОБАВЛЕНИЮ:")
        print("    • Вводите номера в формате: BP<номер> (например: BP26002813)")
        print("    • После каждого номера нажимайте Enter для добавления в очередь")
        print("    • Для завершения ввода оставьте строку пустой и просто нажмите Enter")
        print("-" * 60)

        while True:
            bp_input = input("\n  Введите номер BP (или Enter для завершения): ").strip().upper()

            if bp_input == '':
                break

            # Ручной ввод допускает BP с любым годом (включая < 26) — это override
            # для экстренных технических изменений, которые ещё не попали в отчёт.
            # Автоматический поток BP фильтруется через is_valid_bp_number().
            if bp_input.startswith(BP_FILE_PREFIX) and bp_input[len(BP_FILE_PREFIX):].isdigit():
                bp_filename = f"{bp_input}.xlsx"
                full_bp_path = os.path.join(bp_target_dir, bp_filename)

                if os.path.exists(full_bp_path):
                    if bp_filename not in auto_bp_files and bp_filename not in manual_bp_files:
                        manual_bp_files.append(bp_filename)
                        print(f"    → Файл '{bp_filename}' успешно добавлен в текущую сессию обработки.")
                    else:
                        print(f"    [Внимание] Файл '{bp_filename}' уже находится в списке планирования.")
                else:
                    print(f"    [Ошибка] Файл '{bp_filename}' не обнаружен по целевому пути: '{bp_target_dir}'")
                    print("    Пожалуйста, скачайте его и положите в указанную папку перед вводом.")
            else:
                print(f"    [Ошибка] Неверный формат '{bp_input}'. Используйте инженерный стандарт, например: BP12345")

    # Слияние автоматического списка и ручного ввода
    bp_files = list(set(auto_bp_files + manual_bp_files))

    # Если новых нет и ручной ввод пользователь отклонил (нажал "нет" или оставил пустым)
    if not bp_files:
        print("\n" + "═" * 60)
        print("  Очередь обработки пуста. Новых задач нет.")
        print("  Выполнение программы успешно завершено. До встречи!")
        print("═" * 60)
        wait_for_user()
        sys.exit(0)

    print("\n" + "=" * 60)
    print("  ФИНАЛЬНЫЙ ПЛАН ОБРАБОТКИ ТЕКУЩЕЙ СЕССИИ")
    print("=" * 60)
    for i, f in enumerate(sorted(bp_files), 1):
        print(f"    {i}. {f}")

    wait_for_user("\n  План утвержден. Нажмите Enter для перехода к проверке наличия файлов на диске...")


    # =============================================================================
    # ЭТАП 2: ПРОВЕРКА НАЛИЧИЯ BP ФАЙЛОВ (BP<номер>.xlsx)
    # =============================================================================
    print("\n" + "=" * 60)
    print("ЭТАП 2: ПРОВЕРКА НАЛИЧИЯ BP ФАЙЛОВ (BP<номер>.xlsx)")
    print("=" * 60)

    while True:
        print(f"\n  Сканирование целевой директории: '{bp_target_dir}'...")

        # Фактический список BP-файлов, физически лежащих в директории
        actual_bp_files = set(find_bp_files(
            file_prefix=BP_FILE_PREFIX,
            file_path=bp_target_dir
        ))

        # Ожидаемые файлы: автоматические (из new_bp_set) + добавленные вручную
        expected_bp_files = set(bp_files)

        # Пересечение: ожидаемые, которые реально есть на диске
        existing_files = sorted(expected_bp_files & actual_bp_files)

        # Отсутствующие: ожидаемые, которых нет на диске
        missing_files = sorted(expected_bp_files - actual_bp_files)

        # Лишние: файлы на диске, но не в текущем плане (не из new_bp_set и не введены вручную)
        extra_files = sorted(actual_bp_files - expected_bp_files)

        if extra_files:
            print(f"\n  [Информация] Найдены BP-файлы вне текущего плана ({len(extra_files)}):")
            for f in extra_files:
                print(f"    - {f}")
            print("  Эти файлы НЕ будут обработаны в текущей сессии.")

        # Все запланированные файлы на месте — идеальный путь
        if not missing_files:
            print(f"  [Успех] Все запланированные файлы ({len(existing_files)} шт.) успешно обнаружены.")
            bp_files_to_process = existing_files
            break

        # Часть файлов отсутствует
        print(f"  [Внимание] Обнаружено файлов: {len(existing_files)}. Отсутствует: {len(missing_files)}")
        print("  Список отсутствующих файлов:")
        for f in missing_files:
            print(f"    - {f}")

        print("\n  Варианты действий:")
        print("    Нажмите: 1 - Приостановить программу, докачать эти файлы и повторить проверку")
        if existing_files:
            print("    Нажмите: 2 - Игнорировать отсутствующие и продолжить только с найденными файлами")
        print("    Нажмите: 3 - Отменить всё и выйти из программы")

        # Формируем динамическое приглашение к вводу
        allowed_choices = ['1', '2', '3'] if existing_files else ['1', '3']
        choice_prompt = "  Ваш выбор (1-3): " if existing_files else "  Ваш выбор (1 или 3): "

        while True:
            user_action = ask_user_input(choice_prompt).strip()
            if user_action in allowed_choices:
                break
            print("  Некорректный ввод. Пожалуйста, выберите вариант из списка.")

        if user_action == '1':
            # Даем пользователю шанс докачать файлы прямо сейчас
            print(f"\n  [Пауза] Пожалуйста, скачайте отсутствующие файлы изменений в папку: '{bp_target_dir}'")
            wait_for_user("  Как только файлы будут скопированы, нажмите Enter для повторной проверки...")
            continue  # Возвращаемся на начало цикла while True и сканируем папку заново!

        elif user_action == '2':
            # Продолжаем только с теми, что нашли
            print("\n  [Внимание] Очередь скорректирована. Отсутствующие файлы будут пропущены.")
            bp_files_to_process = existing_files
            break

        elif user_action == '3':
            # Спокойный, штатный выход по требованию пользователя
            print("\n  Выполнение программы отменено пользователем. До встречи!")
            wait_for_user()
            sys.exit(0)

    # Гарантированная финальная проверка перед запуском конвейера
    if not bp_files_to_process:
        print(f"\n  [Ошибка] В директории '{bp_target_dir}' по-прежнему нет ни одного файла для обработки.")
        print("  Конвейер пуст. Программа завершает работу.")
        wait_for_user()
        sys.exit(0)

    print(f"\n  Утверждена итоговая очередь сессии: {len(bp_files_to_process)} файлов к обработке.")
    for i, f in enumerate(bp_files_to_process, 1):
        print(f"    [{i}] {f}")

    wait_for_user("\n  План утвержден. Нажмите Enter для запуска потокового конвейера обработки (ЭТАП 4)...")


    # =============================================================================
    # ЭТАП 3: ПОТОКОВАЯ ОБРАБОТКА BP И СОХРАНЕНИЕ (YYYY-mm-dd_breakpoint_data.xlsx)
    # =============================================================================
    print("\n" + "=" * 60)
    print("ЭТАП 3: ПОТОКОВАЯ ОБРАБОТКА BP И СОХРАНЕНИЕ (YYYY-mm-dd_breakpoint_data.xlsx)")
    print("=" * 60)

    processed_results: Dict[Any, Any] = {}

    for i, bp_file in enumerate(bp_files_to_process, 1):
        print("\n" + "=" * 60)
        print(f"ОБРАБОТКА BP ФАЙЛА {i}/{len(bp_files_to_process)}: {bp_file}")
        print("=" * 60)

        full_bp_input_path = os.path.join(bp_target_dir, bp_file)
        bp_num = bp_file.replace('.xlsx', '')

        print(f"  [Конвейер] Подготовка к пошаговой обработке: '{bp_num}'")
        wait_for_user(f"  Нажмите Enter, чтобы запустить сеанс для {bp_num}...")

        # 1. Запуск интерактивного конвейера для одного файла
        result_dict: Optional[Dict[str, Any]] = process_bp_file(full_bp_input_path)

        # ПРОВЕРКА НА МАРКЕР ПРОПУСКА ФАЙЛА
        if result_dict and result_dict.get('status') == 'SKIP_FILE':
            print(f"  [Конвейер] Файл {bp_file} успешно пропущен по решению оператора.")
            if i < len(bp_files_to_process):
                wait_for_user("\n  Нажмите Enter для перехода к следующему файлу изменений...")
            continue # Переходим к следующей итерации цикла for, обрабатывая следующий файл!

        if result_dict and isinstance(result_dict, dict):
            df_processed = result_dict.get('dataframe')

            if df_processed is not None and not df_processed.empty:
                # Фиксируем в памяти текущей сессии
                processed_results[bp_num] = df_processed

                # 2. ЖЕСТКАЯ ЗАЩИТА АВТОНОМНОГО БЭКАПА (в output_backup_files)
                save_backup(df=df_processed, bp_number=bp_num, backup_root=OUTPUT_BACKUP_ROOT)

                # 3. МОМЕНТАЛЬНОЕ ИНКРЕМЕНТАЛЬНОЕ СОХРАНЕНИЕ НА ДИСК (Транзакция)
                print(f"  [Транзакция] Фиксация изменений {bp_num} в breakpoint_data.xlsx...")

                success_inject = inject_bp_to_global_report(df_processed, bp_num)

                if success_inject:
                    print(f"  [Успех] Прогресс сохранен. {bp_num} добавлен в итоговую базу данных.")
                else:
                    print(f"  [ВНИМАНИЕ] Не удалось выполнить инкрементальное сохранение для {bp_num}!")
                    if not ask_yes_no("  Хотите продолжить обработку следующих файлов, несмотря на ошибку записи?"):
                        print("  Выполнение конвейера остановлено пользователем для исправления проблем с доступом к файлу отчета.")
                        sys.exit(0)

        if i < len(bp_files_to_process):
            wait_for_user("\n  Шаг завершен. Нажмите Enter для перехода к следующему файлу изменений...")


    # =============================================================================
    # ЭТАП 4: ИТОГИ
    # =============================================================================
    print("\n" + "=" * 60)
    print("ЭТАП 4: ИТОГИ")
    print("=" * 60)
    print(f"\nОбработано файлов: {len(processed_results)}/{len(bp_files_to_process)}")

    if processed_results:
        print("\nОбработанные Breakpoint'ы:")
        for bp_number in processed_results:
            print(f"   - {bp_number}")

        # ИТОГОВОЕ ВРЕМЯ РАБОТЫ ПРОГРАММЫ
        total_program_time = time.time() - program_start_time
        print("\n" + "=" * 70)
        print("ИТОГОВОЕ ВРЕМЯ РАБОТЫ ПРОГРАММЫ")
        print("=" * 70)
        print(f"  Общее время: {timedelta(seconds=int(total_program_time))}")
        print(f"  ({(total_program_time/60):.1f} минут)")
        print("=" * 70)

        return processed_results

    print("\nНе обработано ни одного файла.")
    return None


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

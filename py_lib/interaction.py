#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
"""
"""
import sys
from typing import Dict, List, Optional, Tuple

import pandas as pd

from py_lib.etl import str_convert

def _handle_interrupt() -> None:
    """
    Внутренний централизованный обработчик прерывания Ctrl+C.
    
    Выводит интерактивный диалог подтверждения выхода. Изолирует 
    бесконечный цикл проверки намерения пользователя, исключая 
    дублирование блоков except KeyboardInterrupt в интерфейсных функциях.
    """
    print()
    while True:
        try:
            confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
            if confirm == 'да':
                print("\n\nПрограмма прервана пользователем (Ctrl+C)")
                sys.exit(0)
            elif confirm == 'нет':
                print("\nПродолжаем работу...")
                return
            print("Пожалуйста, введите 'да' или 'нет'")
        except (KeyboardInterrupt, EOFError):
            print("\n\nПрограмма экстренно завершена.")
            sys.exit(0)


def ask_yes_no(prompt: str, default_yes: bool = True) -> bool:
    """
    Универсальный запрос для подтверждения операций пользователя (Да/Нет).

    Аргументы:
        prompt (str):       Текст вопроса.
        default_yes (bool): Значение при нажатии Enter без ввода текста.

    Возвращается:
        bool: True если пользователь выбрал 'да', False если 'нет'.
    """
    suffix = " (Да/нет) [Да]: " if default_yes else " (да/Нет) [Нет]: "
    while True:
        try:
            user_input = input(f"{prompt}{suffix}").strip().lower()
            if not user_input:
                return default_yes
            if user_input in ("да", "д", "yes", "y"):
                return True
            if user_input in ("нет", "н", "no", "n"):
                return False
            print("Некорректный ввод. Пожалуйста, введите 'да' или 'нет'.")
        except KeyboardInterrupt:
            _handle_interrupt()
        except EOFError:
            print("\n\nКонец ввода. Программа завершена.")
            sys.exit(0)


def clear_screen():
    """
    Очистка экрана консоли (кросс-платформенная).
    
    Использует ANSI-escape последовательности для очистки экрана.
    Это работает в большинстве современных терминалов.
    
    Поддерживаемые платформы:
        - Windows (с поддержкой ANSI)
        - Linux
        - macOS
        - WSL
    
    Возвращается:
        None
    """
    # ANSI escape последовательность для очистки экрана и сброса курсора в (0,0)
    # \033[2J - очистить экран
    # \033[H - переместить курсор в верхний левый угол
    print('\033[2J\033[H', end='')


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


def wait_for_user(prompt: str = "\nНажмите Enter для продолжения...") -> str:
    """
    Ожидает нажатия клавиши Enter от пользователя.

    Аргументы:
        prompt (str): Текст приглашения к вводу. По умолчанию содержит инструкцию.

    Возвращается:
        str: Введённая пользователем строка.
    """
    while True:
        try:
            return input(prompt)
        except KeyboardInterrupt:
            _handle_interrupt()
        except EOFError:
            print("\n\nКонец ввода. Программа завершена.")
            sys.exit(0)


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


def confirm_step(
    step_name: str,
    df_current: pd.DataFrame,
    saved_state: Optional[pd.DataFrame]
) -> Tuple[bool, pd.DataFrame, Optional[pd.DataFrame]]:
    """
    Запрашивает у пользователя подтверждение после выполнения шага.

    Поддерживает:
        - Enter: подтверждение, сохранение нового состояния
        - 'retry': отмена изменений и восстановление предыдущего состояния

    Аргументы:
        step_name (str):             Имя шага для вывода сообщений.
        df_current (pd.DataFrame):   Текущий DataFrame после выполнения шага.
        saved_state (pd.DataFrame):  Сохранённое состояние перед шагом (может быть None).

    Возвращается:
        Tuple[bool, pd.DataFrame, Optional[pd.DataFrame]]:
            - continue_flag (bool):  True, если нужно продолжить, False, если 'retry'
            - df (pd.DataFrame):     Текущий или восстановленный DataFrame
            - new_saved_state (Optional[pd.DataFrame]): Новое сохранённое состояние
    """
    print(f"\n  Шаг '{step_name}' выполнен.")

    while True:
        try:
            user_input = input(
                "\nПроверьте результат. Если всё корректно, нажмите Enter. Если нужно повторить шаг, введите 'retry': "
            ).strip().lower()

            if user_input == 'retry':
                print(f"Повторяем шаг '{step_name}'...\n")
                restored_df = restore_state(saved_state, step_name)
                if restored_df is not None:
                    return False, restored_df, saved_state
                return False, df_current, saved_state

            print("Продолжаем...\n")
            new_saved_state = save_state_before_step(df_current)
            return True, df_current, new_saved_state

        except KeyboardInterrupt:
            _handle_interrupt()
        except EOFError:
            print("\n\nОбнаружен конец ввода. Программа завершена.")
            sys.exit(0)


def get_bp_status(
    bp_number: str
) -> str:
    """
    Запрашивает у пользователя статус обработки для текущего BP файла.

    Доступные варианты:
    1. Согласован/Approved
    2. Опубликован/Published
    3. Закрыт/Closed
    4. Другое (ввод вручную)

    Аргументы:
        bp_number (str): Номер BP для вывода в сообщении.

    Возвращается:
        str: Выбранный статус (с переводом строки для двустрочного формата).
    """
    print(f"\n  Для BP файла {bp_number} укажите статус обработки:")
    print("  Доступные варианты:")
    print("    1 - Согласован\n        Approved")
    print("    2 - Опубликован\n        Published")
    print("    3 - Закрыт\n        Closed")
    print("    4 - Другое (ввести вручную)")

    while True:
        try:
            choice = input("  Ваш выбор (1-4, или Enter для 'Согласован/Approved'): ").strip()

            if choice == '' or choice == '1':
                return "Согласован\nApproved"
            if choice == '2':
                return "Опубликован\nPublished"
            if choice == '3':
                return "Закрыт\nClosed"
            if choice == '4':
                while True:
                    try:
                        custom_status = input("  Введите свой статус: ").strip()
                        break
                    except KeyboardInterrupt:
                        _handle_interrupt()
                return custom_status if custom_status else "Согласован\nApproved"

            print("  Неверный выбор. Пожалуйста, введите число от 1 до 4.")
        except KeyboardInterrupt:
            _handle_interrupt()
        except EOFError:
            print("\n\nКонец ввода. Программа завершена.")
            sys.exit(0)


def interactive_translation(
    data: List[str],
    field_name: str,
    examples: Optional[Dict[str, str]] = None
) -> Dict[str, str]:
    """
    Интерактивный ввод переводов для списка уникальных значений.

    Аргументы:
        data (list): Список уникальных значений для перевода.
        field_name (str): Название поля для вывода (например, "названий деталей").
        examples (dict, optional): Словарь примеров переводов для отображения.

    Возвращается:
        Dict[str, str]: Словарь соответствий {оригинал: перевод}.
    """
    if not data:
        return {}

    translations: Dict[str, str] = {}

    print(f"\nПеревод {field_name}:")
    print(f"Найдено {len(data)} уникальных значений")

    if examples:
        print("Примеры переводов (можно использовать как шаблон):")
        for ch, ru in examples.items():
            print(f"     {ch[:50]}... → {ru[:50]}...")

    print("\nСписок всех уникальных значений:\n" + "-" * 60)
    for i, value in enumerate(data, 1):
        print(f"  {i}. {value}")
    print("-" * 60 + "\n\nИнструкция:")
    print("  • Введите перевод и нажмите Enter → оригинальный текст будет заменён на перевод")
    print("  • Нажмите Enter без перевода → текст останется оригинальным (без изменений)")

    for i, value in enumerate(data, 1):
        if pd.isna(value) or value == '':
            translations[value] = value
            continue

        print(f"\n[{i}/{len(data)}] Оригинал: {value}")
        while True:
            try:
                user_input = input("Введите перевод (или просто Enter чтобы оставить оригинал): ").strip()
                break
            except KeyboardInterrupt:
                _handle_interrupt()
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


def user_input_for_single_bp(
    df_current: pd.DataFrame,
    bp_number: str
) -> pd.DataFrame:
    """
    Запрашивает у пользователя ввод общих данных (Batch fact, Change Date) для всего BP.

    Аргументы:
        df_current (pd.DataFrame): Текущий DataFrame для обработки.
        bp_number (str):           Номер обрабатываемого бизнес-процесса.

    Возвращается:
        pd.DataFrame: Копия DataFrame с заполненными общими данными.
    """
    print(f"\n--- Ввод данных для {bp_number} ---")
    print(f"Всего строк для обработки: {len(df_current)}")

    df_result = df_current.copy()
    print("\nВведите общие данные для всего технического изменения:")

    # Безопасное извлечение текущих значений с помощью str_convert из etl.py
    has_rows = len(df_result) > 0
    current_batch_fact = str_convert(df_result['Batch fact'].iloc[0]) if has_rows and 'Batch fact' in df_result.columns else ''
    current_change_date = str_convert(df_result['Change Date'].iloc[0]) if has_rows and 'Change Date' in df_result.columns else ''

    # Ввод для Batch fact
    print(f"Текущее Batch fact: {current_batch_fact if current_batch_fact else '(пусто)'}")
    while True:
        try:
            batch_fact_input = input("Введите Batch fact (или Enter, чтобы оставить пустым): ").strip()
            break
        except KeyboardInterrupt:
            _handle_interrupt()

    # Ввод для Change Date
    print(f"Текущее Change Date: {current_change_date if current_change_date else '(пусто)'}")
    while True:
        try:
            change_date_input = input("Введите Change Date (ГГГГ-ММ-ДД или Enter, чтобы оставить пустым): ").strip()
            break
        except KeyboardInterrupt:
            _handle_interrupt()

    # Применение введенных данных к массиву
    if batch_fact_input:
        df_result['Batch fact'] = batch_fact_input
        print(f"  Batch fact '{batch_fact_input}' применён ко всем {len(df_result)} строкам")

    if change_date_input:
        df_result['Change Date'] = change_date_input
        print(f"  Change Date '{change_date_input}' применён ко всем {len(df_result)} строкам")

    if not batch_fact_input and not change_date_input:
        print("  Данные не введены. Будут заполнены позже в Excel.")

    return df_result


def show_dataframe_preview(
    df: Optional[pd.DataFrame],
    step_name: str,
    max_rows: int = 10,
    focus_columns: Optional[List[str]] = None,
    max_colwidth: int = 40
) -> None:
    """
    Отображает первые строки DataFrame для визуального контроля результатов шага.

    Пустые значения (NaN, None) показываются как пустая ячейка.
    Числовые значения отображаются как есть (0 не заменяется).

    Аргументы:
        df (Optional[pd.DataFrame]): Массив данных для анализа (может быть None).
        step_name (str):             Имя текущего шага для вывода заголовка.
        max_rows (int):              Максимальное количество строк для вывода. По умолчанию 10.
        focus_columns (list, opt):   Список целевых колонок для отображения.
        max_colwidth (int):          Максимальная ширина ячейки в символах. По умолчанию 40.
    """
    if df is None or df.empty:
        print(f"\n[Preview после шага: {step_name}]")
        print("  DataFrame пуст или отсутствует!")
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

    # Создаем безопасную копию выбранных колонок для превью
    preview_df = df[display_cols].head(max_rows).copy()

    # Форматируем ячейки под требования консольного вывода
    for col in preview_df.columns:
        # Заменяем пустые значения на пустую строку
        preview_df[col] = preview_df[col].apply(
            lambda x: '' if pd.isna(x) else str(x)
        )
        # Обрезаем длинный текст
        preview_df[col] = preview_df[col].apply(
            lambda x: (x[:max_colwidth] + '…') if len(x) > max_colwidth else x
        )

    # Вывод с использованием контекстного менеджера pandas
    with pd.option_context(
        'display.max_columns', len(display_cols),
        'display.width', None,
        'display.max_colwidth', max_colwidth,
        'display.show_dimensions', False,
        'display.unicode.east_asian_width', True
    ):
        print(preview_df.to_string(index=False, na_rep=''))
    print()

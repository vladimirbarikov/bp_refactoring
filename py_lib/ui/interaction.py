#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
# pylint: disable=import-outside-toplevel
"""
BP Refactoring Tool - Пользовательское взаимодействие и консольный интерфейс (Interface Layer).

Модуль спроектирован как изолированный визуально-интерактивный слой проекта.
Он инкапсулирует логику двустороннего консольного диалога с оператором,
пошагового UI-контроля вычислений и форматированного вывода данных на экран
для ручной валидации.

Архитектура модуля разделена на 3 функциональных блока:

    1. СИСТЕМНЫЕ УТИЛИТЫ И ПЕРЕХВАТ ПРЕРЫВАНИЙ (Signal & Stream Handling)
       Служебные компоненты нижнего уровня для очистки экрана терминала
       и защиты потока ввода от аварийных сигналов (Ctrl+C, EOFError).
       Изолируют диалоги завершения программы, исключая дублирование
       блоков try-except в коде верхних уровней.

    2. БАЗОВЫЙ ИНТЕРАКТИВНЫЙ ВВОД/ВЫВОД (Basic UI Components)
       Универсальные инструменты стандартизированного текстового
       взаимодействия: отрисовка заголовков этапов, валидация простых
       ответов «Да/Нет», безопасный ввод текста и чисел, контролируемые
       паузы в работе конвейера.

    3. КОНТРОЛЬ КОНВЕЙЕРА И ПРЕДВЫБОРКА ДАННЫХ (Pipeline Control & Data Preview)
       Инструменты визуального контроля результатов: создание временных
       снимков DataFrame в оперативной памяти для функции отката шага (retry),
       а также матричный вывод таблиц в консоль с автоподгонкой ширины ячеек.

Экспортируемые функции (по блокам):
    Блок 1: clear_screen
    Блок 2: print_step_header, ask_yes_no, ask_user_input,
            ask_int_or_nan, wait_for_user
    Блок 3: save_state, restore_state, confirm_step,
            show_dataframe_preview

Правила взаимодействия с бизнес-логикой:
    - Модуль является исключительно интерфейсным представлением.
    - Модуль НЕ производит инженерных расчётов, нормализации, парсинга
      Excel-файлов или физической записи данных на диск.
    - Вся собранная от пользователя информация или управляющие команды
      возвращаются наверх в виде базовых структур данных (str, int, bool)
      или копий DataFrame.

Использование:
    from py_lib.ui.interaction import (
        print_step_header,
        wait_for_user,
        confirm_step,
        show_dataframe_preview,
    )

Версия: 1.1
Совместимость: Python 3.14.4+, Pandas 3.0.3+
Поддержка: PLD Engineering Center
Дата создания: 2026-09-10
Дата изменения: 2026-09-11
Лицензия: MIT
Статус: Production
"""

import os
import subprocess
import sys
from typing import List, Optional, Tuple, Union

import numpy as np
import pandas as pd

# ============================================================
# СИСТЕМНЫЕ УТИЛИТЫ И ПЕРЕХВАТ ПРЕРЫВАНИЙ
# ============================================================
def clear_screen() -> None:
    """
    Очистка экрана консоли (кросс-платформенная).
    
    Определяет тип операционной системы и выполняет системную команду
    очистки экрана через безопасный модуль subprocess.
    """
    try:
        if os.name == 'nt':
            subprocess.run(['cls'], shell=True, check=True)
        else:
            subprocess.run(['clear'], check=True)
    except (subprocess.SubprocessError, FileNotFoundError):
        print('\033[2J\033[H', end='')


def _handle_interrupt() -> None:
    """
    Внутренний централизованный обработчик прерывания Ctrl+C и EOFError.

    При срабатывании аварийного сигнала:
        1. Спрашивает у оператора подтверждение завершения программы.
        2. При ответе «да» — завершает программу через sys.exit(0).
        3. При ответе «нет» — возвращает управление вызывающему коду,
           чтобы тот продолжил цикл ввода с того же места.
        4. При повторном аварийном сигнале в момент диалога —
           завершает программу экстренно.
    """
    print()
    while True:
        try:
            confirm = input("\nВы действительно хотите прекратить работу программы (да/нет): ").strip().lower()
            if confirm in ('да', 'д', 'yes', 'y'):
                print("\nРабота программы прекращена пользователем.")
                sys.exit(0)
            elif confirm in ('нет', 'н', 'no', 'n'):
                print("\nВозврат к выполнению задачи...")
                return
            print("Пожалуйста, введите 'да' или 'нет'.")
        except (KeyboardInterrupt, EOFError):
            print("\n\nПрограмма экстренно завершена.")
            sys.exit(0)


# ============================================================
# БАЗОВЫЙ ИНТЕРАКТИВНЫЙ ВВОД/ВЫВОД
# ============================================================
def print_step_header(step_num: int, total_steps: int, description: str) -> None:
    """
    Выводит форматированный заголовок шага обработки.

    Формат вывода:
        ============================================================
        ШАГ {step_num}/{total_steps}: {description}
        ============================================================

    Используется в пайплайне process_bp_file() для визуального
    разделения 25 шагов обработки одного BP-файла.

    Аргументы:
        step_num (int):    Номер текущего шага.
        total_steps (int): Общее количество шагов в пайплайне.
        description (str): Текстовое описание шага.

    Возвращается:
        None
    """
    width = 60
    print("\n" + "=" * width)
    print(f"ШАГ {step_num}/{total_steps}: {description}")
    print("=" * width)


def ask_yes_no(prompt: str, default_yes: bool = True) -> bool:
    """
    Универсальный запрос для подтверждения операций пользователя (Да/Нет).

    Поведение:
        - Пустой ввод (просто Enter) → возвращает значение default_yes.
        - «да», «д», «yes», «y»      → возвращает True.
        - «нет», «н», «no», «n»      → возвращает False.
        - Любой другой ввод          → повторяет запрос.
        - Ctrl+C                     → вызывает _handle_interrupt().

    Аргументы:
        prompt (str):         Текст вопроса оператору.
        default_yes (bool):   Значение по умолчанию при пустом вводе.
                              По умолчанию True.

    Возвращается:
        bool: True при согласии, False при отказе.
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
            continue
        except EOFError:
            print("\n\nКонец ввода. Программа завершена.")
            sys.exit(0)

def ask_user_input(prompt: str) -> str:
    """
    Универсальный защищённый ввод текста с перехватом Ctrl+C.

    Возвращает строку без ведущих/замыкающих пробелов (выполняет .strip()).
    При случайном нажатии Ctrl+C вызывает диалог подтверждения выхода
    через _handle_interrupt():
        - Если оператор отказывается выходить — возврат к строке ввода.
        - Если соглашается — программа завершается через sys.exit(0).

    Аргументы:
        prompt (str): Текст приглашения к вводу.

    Возвращается:
        str: Введённая строка без ведущих и замыкающих пробелов.
    """
    while True:
        try:
            return input(prompt).strip()
        except KeyboardInterrupt:
            _handle_interrupt()  # Показывает диалог "Вы действительно хотите выйти?"
            # Если пользователь выбрал "нет", цикл уйдет на новый круг и заново покажет строку ввода
            continue


def ask_int_or_nan(
    prompt: str
) -> Union[int, float]:
    """
    Запрашивает у пользователя целое число ≥ 1.

    Поведение:
        - Пустой ввод (просто Enter) → возвращает np.nan (пропуск).
        - Целое число ≥ 1            → возвращает int.
        - Некорректный ввод          → повторяет запрос.

    Аргументы:
        prompt (str): Текст приглашения к вводу.

    Возвращается:
        Union[int, float]: Целое число ≥ 1 или np.nan при пропуске.
    """
    while True:
        raw = ask_user_input(prompt).strip()
        if not raw or raw == ' ':
            return np.nan
        if raw.isdigit() and int(raw) >= 1:
            return int(raw)
        print("  [Ошибка] Введите целое число ≥ 1 или нажмите пробел для пропуска.")


def wait_for_user(prompt: str = "\nНажмите Enter для продолжения...") -> str:
    """
    Ожидает нажатия клавиши Enter от пользователя (контролируемая пауза).

    Применяется для остановки конвейера в ключевых точках, чтобы оператор
    успел прочитать выведенный результат перед продолжением работы.

    Поведение при аварийных сигналах:
        - Ctrl+C    → вызывает _handle_interrupt() и продолжает ожидание.
        - EOFError  → программа завершается через sys.exit(0).

    Аргументы:
        prompt (str): Текст приглашения. По умолчанию «Нажмите Enter для продолжения...».

    Возвращается:
        str: Введённая пользователем строка (обычно пустая).
    """
    while True:
        try:
            return input(prompt)
        except KeyboardInterrupt:
            _handle_interrupt()
            continue
        except EOFError:
            print("\n\nКонец ввода. Программа завершена.")
            sys.exit(0)


# ============================================================
# КОНТРОЛЬ КОНВЕЙЕРА И ПРЕДВЫБОРКА ДАННЫХ
# ============================================================
def save_state(df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    """
    Сохраняет глубокую копию DataFrame для возможности отката шага (retry).

    Применяется в начале каждого шага пайплайна: если оператор решит
    повторить шаг, состояние можно восстановить через restore_state().

    Аргументы:
        df (Optional[pd.DataFrame]): DataFrame для сохранения.
                                     Может быть None — тогда вернётся None.

    Возвращается:
        Optional[pd.DataFrame]: Глубокая копия DataFrame или None,
                                если на входе был None.
    """
    if df is not None:
        print("  [Сохранено состояние перед шагом]")
        return df.copy(deep=True)
    return None


def restore_state(saved_df: Optional[pd.DataFrame], step_name: str) -> Optional[pd.DataFrame]:
    """
    Восстанавливает сохранённое состояние DataFrame для повтора шага.

    Применяется внутри confirm_step() при выборе пользователем режима 'retry'.

    Аргументы:
        saved_df (Optional[pd.DataFrame]): Сохранённый ранее DataFrame.
        step_name (str):                   Имя шага для вывода в сообщении.

    Возвращается:
        Optional[pd.DataFrame]: Восстановленная копия DataFrame
                                или None при отсутствии сохранённого состояния.
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

    Поведение:
        - Enter        → шаг фиксируется, возврат нового состояния.
        - «retry»      → шаг повторяется, возврат сохранённого состояния.
        - Любой ввод   → трактуется как подтверждение (не 'retry').
        - Ctrl+C       → вызывает _handle_interrupt().

    Аргументы:
        step_name (str):                      Имя шага для вывода в сообщении.
        df_current (pd.DataFrame):            Текущий DataFrame после выполнения шага.
        saved_state (Optional[pd.DataFrame]): Сохранённое состояние ДО шага.

    Возвращается:
        Tuple[bool, pd.DataFrame, Optional[pd.DataFrame]]:
            - continue_flag (bool):  True если нужно продолжить, False при retry.
            - df (pd.DataFrame):     Текущий или восстановленный DataFrame.
            - new_saved_state:       Новое сохранённое состояние.
    """
    print(f"\n  Шаг '{step_name}' выполнен.")
    while True:
        try:
            user_input = ask_user_input(
                "\nПроверьте результат. Если всё корректно, нажмите Enter. Если нужно повторить шаг, введите 'retry': "
            ).strip().lower()

            if user_input == 'retry':
                print(f"Повторяем шаг '{step_name}'...\n")
                restored_df = restore_state(saved_state, step_name)
                if restored_df is not None:
                    return False, restored_df, saved_state
                return False, df_current, saved_state

            print("Продолжаем...\n")
            new_saved_state = save_state(df_current)
            return True, df_current, new_saved_state

        except KeyboardInterrupt:
            _handle_interrupt()
            continue
        except EOFError:
            print("\n\nОбнаружен конец ввода. Программа завершена.")
            sys.exit(0)


def show_dataframe_preview(
    df: Optional[pd.DataFrame],
    step_name: str,
    max_rows: int = 10,
    focus_columns: Optional[List[str]] = None,
    max_colwidth: int = 30
) -> None:
    """
    Отображает первые N строк DataFrame для визуального контроля результатов шага.

    Если focus_columns не задан — показывает первые 5 колонок.
    Если задан — показывает только те колонки из focus_columns,
    которые реально присутствуют в df.

    Пустые значения (NaN, None, NaT) выводятся как пустые ячейки.
    Длинные строковые значения обрезаются до max_colwidth с '...' в конце.

    Аргументы:
        df (Optional[pd.DataFrame]):  DataFrame для отображения (может быть None).
        step_name (str):              Название шага для заголовка.
        max_rows (int):               Количество строк для показа. По умолчанию 10.
        focus_columns (Optional[List[str]]): Список колонок для приоритетного показа.
        max_colwidth (int):           Максимальная ширина ячейки. По умолчанию 30.
    """
    if df is None or df.empty:
        print(f"\n[Preview после шага: {step_name}]")
        print("  DataFrame пуст или отсутствует!")
        return

    print(f"\n[Preview после шага: {step_name}]")

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

    preview_df = df[display_cols].head(max_rows).copy()
    for col in preview_df.columns:
        preview_df[col] = preview_df[col].apply(
            lambda x: '' if pd.isna(x) else str(x)
        )
        preview_df[col] = preview_df[col].apply(
            lambda x: (x[:max_colwidth] + '...') if len(x) > max_colwidth else x
        )

    with pd.option_context(
        'display.max_columns', len(display_cols),
        'display.width', 120,
        'display.max_colwidth', max_colwidth,
        'display.show_dimensions', False,
    ):
        print(preview_df.to_string(index=False, na_rep=''))
    print()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pylint: disable=line-too-long
# pylint: disable=too-many-lines
"""
Модуль пользовательского взаимодействия и консольного интерфейса (Interface Layer).

Данный модуль спроектирован как изолированный визуально-интерактивный слой проекта 
BREAKPOINT_REFACTORING_TOOL. Он инкапсулирует в себе логику двустороннего консольного 
диалога с оператором, пошагового UI-контроля вычислений и форматированного вывода 
данных на экран для ручной валидации.

Архитектура модуля разделена на 3 функциональных блока:
1. СИСТЕМНЫЕ УТИЛИТЫ И ПЕРЕХВАТ ПРЕРЫВАНИЙ (Signal & Stream Handling)
   Служебные компоненты нижнего уровня для очистки экрана терминала и защиты потока 
   ввода от аварийных сигналов (Ctrl+C, EOFError). Изолируют диалоги завершения программы, 
   исключая дублирование блоков try-except в коде верхних уровней.
2. БАЗОВЫЙ ИНТЕРАКТИВНЫЙ ВВОД/ВЫВОД (Basic UI Components)
   Универсальные инструменты для стандартизированного текстового взаимодействия: 
   отрисовка заголовков текущих этапов, валидация простых ответов «Да/Нет» и 
   организация контролируемых пауз в работе конвейера.
3. КОНТРОЛЬ КОНВЕЙЕРА И ПРЕДВЫБОРКА ДАННЫХ (Pipeline Control & Data Preview)
   Инструменты визуального контроля результатов: создание временных снимков DataFrame 
   в оперативной памяти для обеспечения функции отката шага (retry), а также матричный 
   вывод таблиц в консоль с автоподгонкой ширины ячеек.

Правила взаимодействия с бизнес-логикой:
    - Модуль является исключительно интерфейсным представлением.
    - Модуль НЕ производит инженерных расчетов, нормализации, парсинга Excel-файлов 
      или физической записи данных на диск.
    - Вся собранная от пользователя информация или управляющие команды возвращаются 
      наверх в виде базовых структур данных (str, dict, bool) или копий DataFrame.

Версия: 1.1
Совместимость: Python 3.14.4, Pandas 3.0.3
Поддержка: PLD Engineering Center
Дата создания: 2026-09-10
Дата изменения: 2026-09-11
Лицензия: MIT
Статус: Production
"""
import os
import subprocess
import sys
from typing import List, Optional, Tuple

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
    Внутренний централизованный обработчик прерывания Ctrl+C.
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
    """
    width = 60
    print("\n" + "=" * width)
    print(f"ШАГ {step_num}/{total_steps}: {description}")
    print("=" * width)


def ask_yes_no(prompt: str, default_yes: bool = True) -> bool:
    """
    Универсальный запрос для подтверждения операций пользователя (Да/Нет).
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


def wait_for_user(prompt: str = "\nНажмите Enter для продолжения...") -> str:
    """
    Ожидает нажатия клавиши Enter от пользователя.
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
    Сохраняет состояние DataFrame перед выполнением шага для возможности отката.
    """
    if df is not None:
        print("  [Сохранено состояние перед шагом]")
        return df.copy(deep=True)
    return None


def restore_state(saved_df: Optional[pd.DataFrame], step_name: str) -> Optional[pd.DataFrame]:
    """
    Восстанавливает сохранённое состояние DataFrame.
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
    Отображает первые строки DataFrame для визуального контроля результатов шага.
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

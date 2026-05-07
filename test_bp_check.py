import os
from typing import Optional
from datetime import datetime

import pandas as pd

def clear_screen():
    """Очистка экрана консоли"""
    os.system('cls' if os.name == 'nt' else 'clear')

def safe_str_convert(
        value,
        default=''
    ) -> str:
    """
    Безопасное преобразование значения в строку
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
    Безопасное преобразование значения в дату
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


def show_dataframe_preview(
        df, max_rows=5,
        focus_columns=None,
        max_colwidth=40
    ) -> Optional[pd.DataFrame]:
    """
    Отображает первые строки DataFrame для визуального контроля
    max_rows: количество строк для отображения
    focus_columns: список колонок для отображения (если None - показывает первые 5)
    max_colwidth: максимальная ширина содержимого колонки в символах
    """

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


# Чтение данных
df_bp_list = pd.read_excel('bp_list_2025-2026.xlsx', sheet_name='BP', header=0)
df_bp_list_new = df_bp_list.dropna(how='all').copy()

# ПРЕОБРАЗОВАНИЕ КОЛОНОК перед фильтрацией
df_bp_list_new['BP'] = df_bp_list_new['BP'].apply(safe_str_convert)
df_bp_list_new['IsUnBomBP'] = df_bp_list_new['IsUnBomBP'].apply(safe_str_convert)
df_bp_list_new['Status'] = df_bp_list_new['Status'].apply(safe_str_convert)
df_bp_list_new['Date'] = df_bp_list_new['Date'].apply(lambda x: safe_date_convert(x, pattern='%Y-%m-%d'))
df_bp_list_new['Part Name (E)'] = df_bp_list_new['Part Name (E)'].apply(safe_str_convert)

# Применение фильтров
df_bp_list_filtered = df_bp_list_new[
    (df_bp_list_new['IsUnBomBP'] == 'No') &
    (df_bp_list_new['Status'] != 'Closed') &
    ((df_bp_list_new['Date'] >= datetime(2026, 4, 1)) | (df_bp_list_new['Date'].isna())) &
    # Фильтр: Part Name (E) НЕ содержит слово software (регистронезависимо)
    (~df_bp_list_new['Part Name (E)'].str.lower().str.contains('software', na=False))
].copy()

# Чтение остальных данных
df_bp_data = pd.read_excel('2026-05-05_breakpoint_data.xlsx', sheet_name='pivot', header=2)
df_bp_data_new = df_bp_data.dropna(how='all').copy()

# ПРЕОБРАЗОВАНИЕ КОЛОНОК
df_bp_data_new['BP_No'] = df_bp_data_new['BP_No'].apply(safe_str_convert)
df_bp_data_new['Status'] = df_bp_data_new['Status'].apply(safe_str_convert)

clear_screen()

# Вывод статистики фильтрации
print("=" * 60)
print("СТАТИСТИКА ФИЛЬТРАЦИИ")
print("=" * 60)
print(f"Исходное количество строк: {len(df_bp_list_new)}")
print(f"После всех фильтров: {len(df_bp_list_filtered)}")
print("=" * 60)
print("\n")

# Создание set'ов и вычисление разницы
bp_list_set = set(df_bp_list_filtered['BP'].dropna().astype(str))
bp_data_set = set(df_bp_data_new['BP_No'].dropna().astype(str))
diff_set = bp_list_set - bp_data_set
bp_diff_list = list(diff_set)

# Просмотр результатов
print("=" * 60)
print("НОВЫЕ BP (отсутствуют в breakpoint_data)")
print("=" * 60)
print(f"Всего BP после фильтрации: {len(bp_list_set)}")
print(f"BP в breakpoint_data: {len(bp_data_set)}")
print(f"Новых BP для скачивания: {len(bp_diff_list)}")
print("=" * 60)

if bp_diff_list:
    print("\nСписок BP для скачивания из G-BOM:")
    for i, bp in enumerate(sorted(bp_diff_list), 1):
        print(f"  {i}. {bp}")
else:
    print("\nНет новых BP для скачивания.")

print("\n" + "=" * 60)

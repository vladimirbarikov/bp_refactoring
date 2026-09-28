# Changelog

All notable changes to the **BREAKPOINT REFACTORING TOOL** project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

Changes currently in development that have not yet been released.

---

## [2.0.1] — 2026-09-28

Patch release. Fixes the application of formatting to the final
`breakpoint_data.xlsx` file and makes the formatting configuration
self-documenting.

### Added

- **`main.py`: automatic check and creation of working directories.**
  New function `ensure_project_directories()` runs at startup before the
  banner and walks through all paths from `core.py` (5 input + 2 output),
  creating missing ones recursively via `os.makedirs()`. Handles four
  specific error kinds (`PermissionError`, `FileExistsError`,
  `NotADirectoryError`, `OSError`) with distinct messages and collects
  failed directories into a separate report. Previously, cloning the
  repository without `input_files/` caused a cryptic file-read error;
  now the tool prepares its own working structure.

### Fixed

- **"FOR WAREHOUSE" zones are now drawn above the correct columns.**
  Previously, the merged headers and color highlighting were applied to
  hard-coded legacy indexes (`range(4, 16)` and `[28, 29]`) that did not
  match the current column order in `BP_DATA_COLUMNS_ORDER`. Zones are now
  defined by column names via the `EXCEL_WAREHOUSE_COLUMNS` constant in
  `core.py`.

- **Removed dead code `columns_wrap_indices = [34, 35, 38]`.**
  The list of indexes had no logical meaning: it applied wrap formatting
  to three numeric columns that did not require it. The list has been
  removed together with the corresponding branch in
  `save_excel_with_formatting()`.

- **On formatting failure the file is no longer silently saved "raw".**
  Previously, on any error (`ImportError` for `xlsxwriter`,
  `PermissionError`, any exception), the function silently degraded to
  `df.to_excel()` without the three-row header and colors. Now the file
  is **not written at all**: the caller receives `False`, the operator
  sees an explicit error, and the BP backup (without formatting) is saved
  separately, as originally intended.

### Changed

- **`py_lib/config/core.py`:**
  - Renamed `OUTPUT_BACKUP_ROOT` → `OUTPUT_BACKUP_DIR`,
    `OUTPUT_BREAKPOINT_DATA_ROOT` → `OUTPUT_BREAKPOINT_DATA_DIR` for
    naming consistency across the project.
  - Added constant `EXCEL_WAREHOUSE_COLUMNS: List[List[str]]` — two
    "FOR WAREHOUSE" zones defined by column names.
  - `EXCEL_COLUMN_WIDTHS` migrated from `Dict[int, int]` to
    `Dict[str, int]` — column widths are now defined by column names
    instead of indexes.
  - Removed the constant for wrap columns (dead code).
  - Module version: `1.0` → `1.1`.

- **`py_lib/engine/etl.py`:**
  - Function `save_excel_with_formatting()` is now **fully autonomous**:
    it no longer imports `py_lib.config.core`; all formatting rules are
    passed through arguments (`theme_colors`, `column_widths`,
    `warehouse_columns`).
  - Parameter `column_width` renamed to `column_widths`, type changed
    from `Dict[int, int]` to `Dict[str, int]`.
  - Added support for the "no formatting" mode: if `theme_colors=None`,
    the file is saved as a plain Excel (no colors, borders, merged
    headers, or bold font).
  - Six warehouse formats collapsed into two (`warehouse_format`,
    `wrap_warehouse_format`).
  - Module version: `1.0` → `1.1`.

- **`main.py`:**
  - Added call to `ensure_project_directories()` at the start of `main()`.
  - In `inject_bp_to_global_report()`, added explicit passing of three
    formatting constants (`EXCEL_THEME_COLORS`, `EXCEL_COLUMN_WIDTHS`,
    `EXCEL_WAREHOUSE_COLUMNS`) to `save_processed_dataframe()`.
  - Product version: `2.0` → `2.0.1`.

### Not changed

- `py_lib/pipeline/processing.py` — single-BP pipeline (25 steps).
- `py_lib/pipeline/find_pairs.py` — Before/After pair matching.
- `py_lib/ui/interaction.py` — console UI.
- `BREAKPOINT_REFACTORING_TOOL.spec` — PyInstaller spec.
- `.github/workflows/build-windows.yml` — CI pipeline.
- `requirements.txt` / `requirements-dev.txt` — dependency set.

### Compatibility

- Python: 3.14.4+
- Pandas: 3.0.3+
- OpenPyXL: 3.1.5+
- XlsxWriter: 3.2.9+

Changes to the public function signatures in `etl.py` are **backward
compatible**: new parameters `excel_theme_colors`, `excel_column_widths`,
and `excel_warehouse_columns` in `save_processed_dataframe()` are declared
with a default value of `None`.

---

## [2.0.0] — 2026-09-24

First public release of BREAKPOINT REFACTORING TOOL V2.0.

### Added

- Full 25-step pipeline for interactive processing of a technical change
  (Breakpoint) with rollback (retry) support at every step.
- Four Before/After pair matching scenarios in `find_pairs.py`.
- ETL engine `py_lib/engine/etl.py`: safe type converters, Unicode
  surrogate cleaning, universal DataFrame normalization, Excel reading
  and writing.
- Centralized configuration module `py_lib/config/core.py`: paths, step
  descriptions, column specifications, EN → RUS translation dictionaries,
  workshop and model matrices, Excel palette and column widths.
- Automatic audit of the incoming G-BOM report and the accumulated
  `breakpoint_data` database to detect new BPs.
- Incremental transactional write of results after processing each BP.
- Standalone backup of each processed BP.
- Formatting of the final `breakpoint_data.xlsx` file: three-row header
  (EN + RUS), merged "FOR WAREHOUSE" cells, color highlighting.
- CI pipeline `.github/workflows/build-windows.yml` for building the
  Windows executable via PyInstaller and automatically creating a GitHub
  Release from a `v*.*.*` tag.

---

[Unreleased]: https://github.com/vladimirbarikov/bp_refactoring/compare/v2.0.1...HEAD
[2.0.1]: https://github.com/vladimirbarikov/bp_refactoring/compare/v2.0.0...v2.0.1
[2.0.0]: https://github.com/vladimirbarikov/bp_refactoring/releases/tag/v2.0.0
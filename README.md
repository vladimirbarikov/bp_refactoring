<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body>

<h1 align="center">
    <br>
    BREAKPOINT REFACTORING TOOL
    <br>
    <img src="https://img.shields.io/badge/version-2.0-blue.svg" alt="Version 2.0">
    <img src="https://img.shields.io/badge/python-3.14.4%2B-green.svg" alt="Python 3.14.4+">
    <img src="https://img.shields.io/badge/pandas-3.0.3%2B-green.svg" alt="Pandas 3.0.3+">
    <img src="https://img.shields.io/badge/license-MIT-green.svg" alt="License MIT">
    <img src="https://img.shields.io/badge/status-production-brightgreen.svg" alt="Status Production">
</h1>

<p align="center">
    <strong>Автоматизированная обработка технических изменений (Breakpoint) в производственных системах</strong><br>
    Разработано для PLD Engineering Center
</p>

<p align="center">
    <a href="#описание">Описание</a> •
    <a href="#архитектура-проекта">Архитектура</a> •
    <a href="#основные-возможности">Возможности</a> •
    <a href="#требования-к-входным-файлам">Требования</a> •
    <a href="#установка-и-запуск">Установка</a> •
    <a href="#инструкция-по-использованию">Инструкция</a> •
    <a href="#контакты-поддержки">Контакты</a>
</p>

<hr>

<h2>📋 Описание</h2>

<p><strong>BREAKPOINT REFACTORING TOOL</strong> — это автоматизированное приложение для обработки технических изменений (Breakpoint) в автомобильной промышленности. Инструмент предназначен для инженеров PLD Engineering Center и позволяет эффективно управлять процессом замены деталей.</p>

<p>Программа выполняет полный цикл обработки данных:</p>
<ul>
    <li>Загрузка и обработка Excel файлов BP (Breakpoint) из системы G-BOM</li>
    <li>Пошаговое интерактивное редактирование данных с возможностью отката изменений</li>
    <li>Перевод названий деталей, поставщиков, описаний и решений</li>
    <li>Обработка цехов, рабочих станций и локализации поставщиков</li>
    <li>Формирование сводной таблицы с разделением деталей на пары «До / После изменения»</li>
    <li>Расчёт количества партий на складе Safety Stock</li>
    <li>Загрузка и извлечение данных из упаковочных листов (18 упаковочных полей)</li>
    <li>Сохранение результата в отформатированный Excel файл для рассылки в rLink</li>
</ul>

<hr>

<h2>🏗️ Архитектура проекта</h2>

<p>Проект построен по слоистой архитектуре: точка входа в корне, вся бизнес-логика — в библиотеке <code>py_lib/</code>.</p>

<pre><code>.
├── main.py                                # Точка входа + оркестрация 4 этапов
├── py_lib/
│   ├── config/
│   │   └── core.py                        # Все константы, справочники, матрицы
│   ├── engine/
│   │   └── etl.py                         # ETL-движок: чтение, очистка, сохранение
│   ├── pipeline/
│   │   ├── find_pairs.py                  # Сопоставление пар Before/After
│   │   └── processing.py                  # Пайплайн одного BP (25 шагов)
│   └── ui/
│       └── interaction.py                 # Консольный UI, ввод/вывод, retry
├── input_files/
│   ├── input_breakpoint_data_files/       # История обработки (breakpoint_data)
│   ├── input_breakpoint_files/            # Входящие BP-файлы
│   ├── input_breakpoint_report_files/     # Отчёты G-BOM (breakpoint_report)
│   ├── input_configuration_files/         # Конфигурация (configuration)
│   └── input_packing_list_files/          # Упаковочные листы
└── output_files/
    ├── output_backup_files/               # Резервные копии каждого BP
    └── output_breakpoint_data_files/      # Итоговый файл breakpoint_data
</code></pre>

<h3>Роли модулей</h3>
<table>
    <thead>
        <tr><th>Модуль</th><th>Назначение</th></tr>
    </thead>
    <tbody>
        <tr><td><code>main.py</code></td><td>Оркестрация сессии: аудит BP, проверка файлов, потоковая обработка, итоги</td></tr>
        <tr><td><code>py_lib/pipeline/processing.py</code></td><td>Пайплайн обработки одного BP: 25 шагов, интерактив, retry</td></tr>
        <tr><td><code>py_lib/pipeline/find_pairs.py</code></td><td>Сопоставление Before/After: 4 сценария связывания пар деталей</td></tr>
        <tr><td><code>py_lib/engine/etl.py</code></td><td>ETL: safe-конвертеры, очистка строк, чтение и сохранение Excel</td></tr>
        <tr><td><code>py_lib/config/core.py</code></td><td>Единое хранилище констант: пути, колонки, матрицы, стили</td></tr>
        <tr><td><code>py_lib/ui/interaction.py</code></td><td>Консольный UI: ввод, подтверждения, retry, превью DataFrame</td></tr>
    </tbody>
</table>

<hr>

<h2>🚀 Основные возможности</h2>

<h3>Точка входа (<code>main.py</code>)</h3>
<ul>
    <li>✅ Автоматическая проверка новых BP в системе G-BOM</li>
    <li>✅ Сохранение истории обработки между запусками</li>
    <li>✅ Автоматическая загрузка последнего обработанного файла</li>
    <li>✅ Форматирование выходного Excel файла (три строки заголовков, цветовое выделение)</li>
</ul>

<h3>Пайплайн обработки BP (<code>py_lib/pipeline/processing.py</code>)</h3>
<p>Выполняет <strong>25 последовательных шагов</strong> обработки одного BP:</p>
<ol>
    <li>Загрузка, аудит и нормализация данных</li>
    <li>Разделение деталей на старые и новые (Before / After)</li>
    <li>Определение статуса технического изменения</li>
    <li>Определение названий моделей по кодам</li>
    <li>Ввод фактической партии (Batch Fact)</li>
    <li>Ввод фактической даты (Change Date)</li>
    <li>Ввод количества Before-деталей в Safety Stock</li>
    <li>Перевод названий деталей</li>
    <li>Перевод официальных названий поставщиков</li>
    <li>Ввод статуса локализации поставщиков</li>
    <li>Ввод данных о местоположении Before-поставщиков</li>
    <li>Ввод данных о местоположении After-поставщиков</li>
    <li>Перевод описания к изменению</li>
    <li>Перевод решения к изменению</li>
    <li>Перевод цвета</li>
    <li>Обработка цехов и рабочих станций</li>
    <li>Перевод требований по утилизации старых деталей</li>
    <li>Перевод требований по взаимозаменяемости</li>
    <li>Загрузка конфигурационного файла</li>
    <li>Расчёт партий в Safety Stock</li>
    <li>Поиск конфигурации для утилизации Before-деталей</li>
    <li>Поиск упаковочных данных для Before-деталей</li>
    <li>Поиск упаковочных данных для After-деталей</li>
    <li>Упорядочивание колонок</li>
    <li>Сохранение результата</li>
</ol>

<h3>Сопоставление пар (<code>py_lib/pipeline/find_pairs.py</code>)</h3>
<ul>
    <li>🔗 Связывание Before/After по <code>Part No.</code></li>
    <li>🔗 Связывание Before/After по <code>Part Name(CHN)</code> (когда номер изменился, а название осталось)</li>
    <li>🔄 Обработка <code>Update Type</code>: Add + Delete → Before/After</li>
    <li>🔄 Обработка <code>Update Type</code>: Add + Replace/Update → Before/After</li>
    <li>📑 Сборка матричных строк по спецификации <code>BP_DATA_COLUMNS_ORDER</code></li>
</ul>

<h3>ETL-движок (<code>py_lib/engine/etl.py</code>)</h3>
<p>Низкоуровневые инструменты Extract / Transform / Load, изолированные от бизнес-констант:</p>
<ul>
    <li>📥 Чтение Excel-файлов с фильтрацией колонок и очисткой от суррогатов Unicode</li>
    <li>🔧 Safe-конвертеры типов: <code>str_convert</code>, <code>int_convert</code>, <code>float_convert</code>, <code>date_convert</code></li>
    <li>🧹 Очистка строк: удаление суррогатов, схлопывание пробелов, нормализация пустых значений</li>
    <li>⚙️ Универсальная нормализация DataFrame к заданным типам (int64, float64, datetime)</li>
    <li>🔍 Поиск свежих файлов по шаблону <code>YYYY-mm-dd_&lt;prefix&gt;</code> и BP-файлов</li>
    <li>📦 Извлечение упаковочных данных из листов SCM (коробки, паллеты, веса)</li>
    <li>💾 Сохранение Excel с трёхстрочными заголовками, цветовым форматированием и шириной колонок</li>
    <li>🗂️ Резервное копирование и инкрементальная запись в глобальную базу <code>breakpoint_data</code></li>
</ul>

<h3>Конфигурация (<code>py_lib/config/core.py</code>)</h3>
<p>Единое хранилище всех констант проекта, сгруппированных в 13 тематических разделов:</p>
<ul>
    <li>📁 Пути к входным и выходным директориям проекта</li>
    <li>📋 Текстовые описания 25 шагов пайплайна (<code>BP_STEP_DESCRIPTIONS</code>)</li>
    <li>🏷️ Префиксы файлов: <code>BP</code>, <code>breakpoint_data</code>, <code>breakpoint_report</code>, <code>configuration</code></li>
    <li>📊 Спецификация колонок всех входных и выходных Excel-файлов</li>
    <li>🔤 Словарь переводов колонок EN → RUS (<code>BP_DATA_TRANSLATION_COLS</code>)</li>
    <li>🎯 Колонки предпоказа для каждого шага (<code>STEP_FOCUS_COLUMNS_MAP</code>)</li>
    <li>🎨 Настройки цветов и ширины колонок для Excel</li>
    <li>🏭 Матрицы цехов, рабочих станций и моделей</li>
    <li>📚 Справочники для интерактивного ввода</li>
</ul>

<h3>Консольный UI (<code>py_lib/ui/interaction.py</code>)</h3>
<p>Изолированный визуально-интерактивный слой, реализующий консольный диалог с оператором:</p>
<ul>
    <li>🖥️ Кросс-платформенная очистка экрана терминала</li>
    <li>❓ Универсальный ввод «Да/Нет» с поддержкой значений по умолчанию</li>
    <li>⌨️ Защищённый ввод текста и целых чисел (с перехватом <kbd>Ctrl+C</kbd>)</li>
    <li>🛑 Централизованная обработка прерываний и <code>EOFError</code></li>
    <li>📋 Заголовки шагов с визуальным разделением (<code>print_step_header</code>)</li>
    <li>💾 Сохранение и восстановление состояния DataFrame (<code>save_state</code> / <code>restore_state</code>)</li>
    <li>🔁 Диалог подтверждения шага с поддержкой режима <code>retry</code></li>
    <li>📊 Матричный вывод DataFrame в консоль с автоподгонкой ширины ячеек</li>
</ul>

<hr>

<h2>📁 Требования к входным файлам</h2>

<p>Перед запуском программы убедитесь, что в рабочей папке присутствуют следующие файлы:</p>

<table>
    <thead>
        <tr><th>Директория</th><th>Файл</th><th>Описание</th><th>Обязательность</th></tr>
    </thead>
    <tbody>
        <tr>
            <td><code>input_breakpoint_report_files/</code></td>
            <td><code>YYYY-mm-dd_breakpoint_report.xlsx</code></td>
            <td>Входящий отчёт G-BOM со списком активных BP</td>
            <td><strong>Обязателен</strong></td>
        </tr>
        <tr>
            <td><code>input_breakpoint_files/</code></td>
            <td><code>BP&lt;номер&gt;.xlsx</code></td>
            <td>Файлы технических изменений из G-BOM</td>
            <td><strong>Обязателен</strong></td>
        </tr>
        <tr>
            <td><code>input_breakpoint_data_files/</code></td>
            <td><code>YYYY-mm-dd_breakpoint_data.xlsx</code></td>
            <td>Накопленная история обработанных BP</td>
            <td>Создаётся автоматически</td>
        </tr>
        <tr>
            <td><code>input_configuration_files/</code></td>
            <td><code>YYYY-mm-dd_configuration.xlsx</code></td>
            <td>Конфигурационный файл с листом <code>common</code></td>
            <td>Опционально</td>
        </tr>
        <tr>
            <td><code>input_packing_list_files/</code></td>
            <td><code>Упаковочный лист партия &lt;номер&gt;.xlsx</code></td>
            <td>Упаковочные листы для извлечения данных о коробках и паллетах</td>
            <td>Требуются для упаковочных данных</td>
        </tr>
    </tbody>
</table>

<hr>

<h2>🛠️ Установка и запуск</h2>

<h3>Предварительные требования</h3>
<ul>
    <li>Python 3.14.4 или выше</li>
    <li>Pandas 3.0.3+</li>
    <li>OpenPyXL 3.1.5+</li>
    <li>XlsxWriter 3.2.9+ (для форматирования Excel)</li>
</ul>

<h3>Установка зависимостей</h3>
<pre><code>pip install pandas openpyxl xlsxwriter numpy</code></pre>

<h3>Запуск из исходного кода</h3>
<pre><code>python main.py</code></pre>

<h3>Сборка исполняемого файла (PyInstaller)</h3>
<pre><code>pyinstaller BREAKPOINT_REFACTORING_TOOL.spec</code></pre>
<p>Исполняемый файл <code>BREAKPOINT_REFACTORING_TOOL.exe</code> будет создан в папке <code>dist/</code>.</p>

<hr>

<h2>📖 Инструкция по использованию</h2>

<h3>1. Начало работы</h3>
<p>Программа отображает заголовок, инструкцию и структуру каталогов. Нажмите <kbd>Enter</kbd> для продолжения.</p>

<h3>2. ЭТАП 1: Проверка новых BP</h3>
<p>Программа автоматически проверяет наличие новых BP в системе G-BOM, сверяет с накопленной базой и выводит список новых BP для скачивания. При необходимости можно добавить дополнительные BP вручную.</p>

<h3>3. ЭТАП 2: Проверка наличия BP-файлов</h3>
<p>Программа сканирует директорию <code>input_breakpoint_files/</code> и сверяет список ожидаемых файлов с физически присутствующими. При необходимости можно докачать отсутствующие файлы или пропустить их.</p>

<h3>4. ЭТАП 3: Пошаговая обработка каждого BP</h3>
<p>Для каждого BP файла выполняется <strong>25 шагов</strong> обработки. Управление:</p>
<ul>
    <li><kbd>Enter</kbd> — подтвердить корректность изменений и перейти к следующему шагу</li>
    <li><code>retry</code> — отменить изменения текущего шага и выполнить его заново</li>
    <li><kbd>Ctrl+C</kbd> — прервать выполнение программы (с подтверждением)</li>
</ul>

<p>После каждого BP выполняется:</p>
<ul>
    <li>Резервное копирование в <code>output_backup_files/YYYY-mm-dd_processed_bp/BP&lt;номер&gt;/</code></li>
    <li>Инкрементальная транзакционная запись в <code>output_breakpoint_data_files/YYYY-mm-dd_breakpoint_data.xlsx</code></li>
</ul>

<h3>5. ЭТАП 4: Итоги</h3>
<p>Программа выводит сводку обработанных BP и общее время работы.</p>

<h3>6. Структура итогового Excel файла</h3>
<ul>
    <li><strong>Строка 0:</strong> объединённые ячейки «ДЛЯ КЛАДОВЩИКОВ»</li>
    <li><strong>Строка 1:</strong> английские названия колонок</li>
    <li><strong>Строка 2:</strong> русские переводы колонок</li>
    <li><strong>Строка 3+:</strong> данные с цветовым форматированием</li>
</ul>

<h3>7. Резервное копирование</h3>
<p>Для каждого обработанного BP создаётся папка: <code>output_backup_files/YYYY-mm-dd_processed_bp/BP&lt;номер&gt;/</code></p>

<hr>

<h2>📊 Структура выходного файла</h2>

<p>Итоговый файл <code>YYYY-mm-dd_breakpoint_data.xlsx</code> содержит <strong>~90 колонок</strong> с информацией по каждой паре деталей Before/After:</p>

<ul>
    <li>Метаданные BP: номер, статус, партии, даты, модель</li>
    <li>Информация о деталях: номер, название, вес, количество на авто и в ящике</li>
    <li>Упаковка: 18 полей для Before и 18 для After (тип, размеры, штабелирование)</li>
    <li>Поставщики: название, локализация, местоположение (страна, город, улица, здание)</li>
    <li>Производство: цех, рабочая станция, требование по утилизации</li>
    <li>Переводы: описание, решение, цвет</li>
</ul>

<p>Полный список колонок — в модуле <code>py_lib/config/core.py</code>, константа <code>BP_DATA_TRANSLATION_COLS</code>.</p>

<hr>

<h2>⚠️ Возможные ошибки и их решение</h2>

<table>
    <thead>
        <tr><th>Ошибка</th><th>Решение</th></tr>
    </thead>
    <tbody>
        <tr>
            <td><code>FileNotFoundError: BP&lt;номер&gt;.xlsx</code></td>
            <td>Поместите файл в директорию <code>input_breakpoint_files/</code></td>
        </tr>
        <tr>
            <td><code>PermissionError: файл открыт в Excel</code></td>
            <td>Закройте файл и повторите попытку</td>
        </tr>
        <tr>
            <td><code>Excel file format cannot be determined</code></td>
            <td>Убедитесь, что файл имеет расширение .xlsx или .xls</td>
        </tr>
        <tr>
            <td><code>STRUCTURE_MISMATCH</code> при сохранении</td>
            <td>Структура колонок BP не совпадает с историей. Проверьте, что все шаги пайплайна выполнены</td>
        </tr>
        <tr>
            <td><code>UnicodeEncodeError</code> при сохранении</td>
            <td>Программа автоматически очищает проблемные символы</td>
        </tr>
    </tbody>
</table>

<hr>

<h2>👥 Контакты поддержки</h2>

<h3>Доступ к системам:</h3>
<ul>
    <li><strong>Система G-BOM:</strong> Ермолаева Майя (PLD/ED)</li>
    <li><strong>Система SCM:</strong> Федин Антон (PLD/WL)</li>
    <li><strong>rLink чаты:</strong>
    <ul>
        <li><code>Break Points</code> — админ: Бровкина Софья (MD/PM)</li>
        <li><code>Breakpoint PLD Info</code> — админ: Ермолаева Майя (PLD/ED)</li>
    </ul>
</li>
</ul>

<h3>Разработчик:</h3>
<ul>
    <li><strong>Бариков Владимир</strong> (PLD/ED) — rLink</li>
</ul>

<hr>

<h2>📄 Лицензия</h2>
<p>MIT License</p>

<hr>

<p align="center">
    <strong>Версия:</strong> 2.0 &nbsp;|&nbsp;
    <strong>Дата создания:</strong> 2026-05-14 &nbsp;|&nbsp;
    <strong>Дата изменения:</strong> 2026-09-24 &nbsp;|&nbsp;
    <strong>Статус:</strong> Production
</p>

</body>
</html>

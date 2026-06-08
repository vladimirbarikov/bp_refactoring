<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Breakpoint Refactoring Tool</title>
</head>
<body>

<h1 align="center">
    <br>
    BREAKPOINT REFACTORING TOOL
    <br>
    <img src="https://img.shields.io/badge/version-1.0-blue.svg" alt="Version 1.0">
    <img src="https://img.shields.io/badge/python-3.12.3%2B-green.svg" alt="Python 3.12.3+">
    <img src="https://img.shields.io/badge/license-MIT-green.svg" alt="License MIT">
    <img src="https://img.shields.io/badge/status-production-brightgreen.svg" alt="Status Production">
</h1>

<p align="center">
    <strong>Автоматизированная обработка технических изменений (Breakpoint) в производственных системах</strong><br>
    Разработано для PLD Engineering Center
</p>

<p align="center">
    <a href="#описание">Описание</a> •
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
    <li>Формирование сводной таблицы с разделением деталей на пары "До / После изменения"</li>
    <li>Расчёт количества партий на складе Safety Stock</li>
    <li>Загрузка и извлечение данных из упаковочных листов</li>
    <li>Сохранение результата в отформатированный Excel файл для рассылки в DingTalk</li>
</ul>

<hr>

<h2>🚀 Основные возможности</h2>

<h3>Основной модуль (<code>bp_main.py</code>)</h3>
<ul>
    <li>✅ Автоматическая проверка новых BP в системе G-BOM</li>
    <li>✅ Сохранение истории обработки между запусками</li>
    <li>✅ Автоматическая загрузка последнего обработанного файла</li>
    <li>✅ Форматирование выходного Excel файла (три строки заголовков, цветовое выделение)</li>
</ul>

<h3>Модуль обработки BP (<code>bp_processing.py</code>)</h3>
<p>Выполняет <strong>17 последовательных шагов</strong> обработки:</p>
<ol>
    <li>Загрузка BP файла</li>
    <li>Выбор нужных колонок</li>
    <li>Выбор статуса технического изменения</li>
    <li>Ввод количества деталей в Safety Stock</li>
    <li>Перевод названий деталей</li>
    <li>Поиск официальных названий поставщиков</li>
    <li>Ввод статуса локализации поставщиков</li>
    <li>Фильтрация китайских символов в описании и решении</li>
    <li>Перевод описания к изменению</li>
    <li>Перевод решения к изменению</li>
    <li>Обработка цветов и Color Code</li>
    <li>Обработка рабочих центров</li>
    <li>Перевод требований по утилизации старых деталей</li>
    <li>Перевод требований по взаимозаменяемости</li>
    <li>Проверка наличия деталей в BOM</li>
    <li>Упорядочивание колонок</li>
    <li>Сохранение результата</li>
</ol>

<h3>Модуль формирования сводной таблицы (<code>bp_summary.py</code>)</h3>
<ul>
    <li>🔗 Поиск и связывание пар деталей Before/After</li>
    <li>📊 Расчёт количества партий на складе (Quantity batches in SS)</li>
    <li>📦 Загрузка данных из упаковочных листов (коробки, паллеты)</li>
    <li>🔄 Автоматическое копирование данных упаковки для совпадающих деталей</li>
    <li>📑 Формирование итоговой Excel-таблицы с 39+ колонками</li>
</ul>

<hr>

<h2>📁 Требования к входным файлам</h2>

<p>Перед запуском программы убедитесь, что в рабочей папке присутствуют следующие файлы:</p>

<table>
    <thead>
        <tr><th>Файл</th><th>Описание</th><th>Обязательность</th></tr>
    </thead>
    <tbody>
        <tr><td><code>bp_list_2025-2026.xlsx</code></td><td>Список всех BP с фильтрами</td><td><strong>Обязателен</strong></td></tr>
        <tr><td><code>bom.xlsx</code></td><td>Bill of Materials (спецификация деталей)</td><td><strong>Обязателен</strong></td></tr>
        <tr><td><code>configuration.xlsx</code></td><td>Конфигурационный файл с листом 'common'</td><td>Опционально</td></tr>
        <tr><td><code>BP&lt;номер&gt;.xlsx</code></td><td>Файлы технических изменений из G-BOM</td><td>Требуются для обработки</td></tr>
        <tr><td><code>packing_list_&lt;партия&gt;.xlsx</code></td><td>Упаковочные листы</td><td>Требуются для упаковочных данных</td></tr>
        <tr><td><code>ГГГГ-ММ-ДД_breakpoint_data.xlsx</code></td><td>Ранее обработанные данные (история)</td><td>Создаётся автоматически</td></tr>
    </tbody>
</table>

<hr>

<h2>🛠️ Установка и запуск</h2>

<h3>Предварительные требования</h3>
<ul>
    <li>Python 3.12.3 или выше</li>
    <li>Pandas 3.0.2+</li>
    <li>OpenPyXL 3.1.5+</li>
    <li>XlsxWriter</li>
</ul>

<h3>Установка зависимостей</h3>
<pre><code>pip install pandas openpyxl xlsxwriter numpy</code></pre>

<h3>Запуск из исходного кода</h3>
<pre><code>python bp_main.py</code></pre>

<h3>Сборка исполняемого файла (PyInstaller)</h3>
<pre><code>pyinstaller BREAKPOINT_REFACTORING_TOOL.spec</code></pre>
<p>Исполняемый файл <code>BREAKPOINT_REFACTORING_TOOL.exe</code> будет создан в папке <code>dist/</code>.</p>

<hr>

<h2>📖 Инструкция по использованию</h2>

<h3>1. Начало работы</h3>
<p>Программа отображает заголовок, инструкцию и список требований. Нажмите <kbd>Enter</kbd> для продолжения.</p>

<h3>2. Проверка новых BP</h3>
<p>Программа автоматически проверяет наличие новых BP в системе G-BOM, применяет фильтры и выводит список новых BP для скачивания.</p>

<h3>3. Ручной ввод BP (опционально)</h3>
<p>После автоматической проверки можно добавить дополнительные BP номера вручную. Формат ввода: <code>BP26002813</code> (с префиксом BP).</p>

<h3>4. Пошаговая обработка каждого BP файла</h3>
<p>Для каждого BP файла выполняется 17 шагов обработки. Управление:</p>
<ul>
    <li><kbd>Enter</kbd> — подтвердить корректность изменений и перейти к следующему шагу</li>
    <li><code>retry</code> — отменить изменения текущего шага и выполнить его заново</li>
    <li><kbd>Ctrl+C</kbd> — прервать выполнение программы (с подтверждением)</li>
</ul>

<h4>Пример ввода перевода:</h4>
<pre><code>[1/5] Оригинал: 发动机盖板
Введите перевод (или просто Enter чтобы оставить оригинал): Крышка двигателя
  → Заменяем на: Крышка двигателя</code></pre>

<h3>5. Формирование сводной таблицы</h3>
<p>После обработки всех BP файлов программа:</p>
<ul>
    <li>Связывает пары деталей Before/After (по Part No. и Part Name)</li>
    <li>Запрашивает Batch fact и Change Date</li>
    <li>Рассчитывает количество партий на складе</li>
    <li>Загружает упаковочные листы для Before и After деталей</li>
    <li>Автоматически копирует данные упаковки для совпадающих деталей</li>
</ul>

<h3>6. Сохранение результата</h3>
<p>Итоговый файл сохраняется в формате: <code>ГГГГ-ММ-ДД_breakpoint_data.xlsx</code></p>
<p><strong>Структура Excel файла:</strong></p>
<ul>
    <li><strong>Строка 0:</strong> объединённые ячейки "ДЛЯ КЛАДОВЩИКОВ"</li>
    <li><strong>Строка 1:</strong> английские названия колонок</li>
    <li><strong>Строка 2:</strong> русские переводы колонок</li>
    <li><strong>Строка 3+:</strong> данные с цветовым форматированием</li>
</ul>

<h3>7. Резервное копирование</h3>
<p>Для каждого обработанного BP создаётся папка: <code>ГГГГ-ММ-ДД_processed_bp/BP&lt;номер&gt;/</code></p>

<hr>

<h2>📊 Структура выходного файла</h2>

<p>Итоговый файл содержит следующие колонки:</p>

<table>
    <thead>
        <tr><th>Колонка (англ)</th><th>Колонка (рус)</th></tr>
    </thead>
    <tbody>
        <tr><td><code>BP_No</code></td><td>Номер переключения</td></tr>
        <tr><td><code>Status</code></td><td>Статус переключения</td></tr>
        <tr><td><code>Batch plan</code></td><td>Партия по плану</td></tr>
        <tr><td><code>New Part Available Date</code></td><td>Дата выхода новой детали</td></tr>
        <tr><td><code>Batch fact</code></td><td>Партия по факту</td></tr>
        <tr><td><code>Change Date</code></td><td>Дата переключения</td></tr>
        <tr><td><code>BOM Product</code></td><td>Модель</td></tr>
        <tr><td><code>Part No. Before</code></td><td>Номер "старой" детали до переключения</td></tr>
        <tr><td><code>Part Name Before</code></td><td>Название "старой" детали до переключения</td></tr>
        <tr><td><code>Quantity in SS</code></td><td>Количество "старых" деталей на Safety Stock</td></tr>
        <tr><td><code>Quantity batches in SS</code></td><td>Количество партий со "старыми" деталями</td></tr>
        <tr><td><code>Part No. After</code></td><td>Номер "новой" детали после переключения</td></tr>
        <tr><td><code>Part Name After</code></td><td>Название "новой" детали после переключения</td></tr>
        <tr><td><code>Color Code</code></td><td>Код цвета</td></tr>
        <tr><td><code>Color Name (RUS)</code></td><td>Название цвета</td></tr>
        <tr><td><code>Comments</code></td><td>Комментарии</td></tr>
    </tbody>
</table>

<hr>

<h2>⚠️ Возможные ошибки и их решение</h2>

<table>
    <thead>
        <tr><th>Ошибка</th><th>Решение</th></tr>
    </thead>
    <tbody>
        <tr><td><code>FileNotFoundError: bp_list_2025-2026.xlsx</code></td><td>Поместите файл в рабочую папку</td></tr>
        <tr><td><code>PermissionError: файл открыт в Excel</code></td><td>Закройте файл и повторите попытку</td></tr>
        <tr><td><code>Excel file format cannot be determined</code></td><td>Убедитесь, что файл имеет расширение .xlsx или .xls</td></tr>
        <tr><td><code>UnicodeEncodeError</code> при сохранении</td><td>Программа автоматически очищает проблемные символы</td></tr>
    </tbody>
</table>

<hr>

<h2>👥 Контакты поддержки</h2>

<h3>Доступ к системам:</h3>
<ul>
    <li><strong>Система G-BOM:</strong> Бариков Владимир / Ермолаева Майя (PLD/ED)</li>
    <li><strong>Система SCM:</strong> Федин Антон (PLD/WL)</li>
    <li><strong>DingTalk чаты:</strong>
        <ul>
            <li><code>Break Point (BP)</code> — админ: Алексеева Елизавета (MD/PM)</li>
            <li><code>Breakpoint PLD Info</code> — админ: Бариков Владимир (PLD/ED)</li>
        </ul>
    </li>
</ul>

<h3>Разработчик:</h3>
<ul>
    <li><strong>Бариков Владимир</strong> (PLD/ED) — DingTalk</li>
</ul>

<hr>

<h2>📄 Лицензия</h2>
<p>MIT License</p>

<hr>

<p align="center">
    <strong>Версия:</strong> 1.0 &nbsp;|&nbsp;
    <strong>Дата создания:</strong> 2026-05-14 &nbsp;|&nbsp;
    <strong>Статус:</strong> Production
</p>

</body>
</html>

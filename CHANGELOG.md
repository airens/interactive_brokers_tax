# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1] - 2020-03-26

### Added
- Первая версия

## [0.2] - 2020-04-26

### Added (ib.py)
- Парсинг разделов отчета "Deposits & Withdrawals" и "Interest"
- Поддержка валют RUB и EURO
- Расчет дохода по программе повышения доходности

### Added (template.docx)
- В таблицы добавлен столбец "Валюта"
- Перечень прикладываемых документов
- Раздел 2.4 по программе повышения доходности

## [0.3] - 2021-01-17

### Added
- Добавлена поддержка (в виде пропуска таблицы) Strategies Interactive Brokers
- Добавлена поддержка несовпадающих по размерам таблиц дивидентов и уплаченных по ним налогов 

### Fixed
- Добавлена сортировка по дате операций купли\продажи перед обработкой для избежания неправильного расчета (в некоторых случаях) FIFO
- Доплата налога по дивидентам больше не может быть отрицательной
- Если ОС не Windows, то скрипт просто не открывает файл-отчет в конце (раньше выдавал ошибку)

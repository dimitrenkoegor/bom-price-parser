# Findchips BOM

- URL: `https://www.findchips.com/bom`
- Может требовать аккаунт Supplyframe. Если требует — остановиться и
  попросить пользователя войти.

## Процедура (уточнить при первом прогоне)

1. Открыть URL, подождать 4–6 с.
2. Загрузить CSV, сопоставить колонки Part Number / Quantity / Manufacturer.
3. Валюта **USD**; если есть фильтр по типу дистрибьютора — только
   authorized/franchised.
4. Дождаться подбора, экспортировать.

## Формат экспорта

Не верифицирован (заглушка в `SITES`). При BrowserExportFormatError —
добавить синонимы в `SITES['findchips']['aliases']` (текст ошибки подскажет),
копию экспорта — в `final/browser/samples/`.

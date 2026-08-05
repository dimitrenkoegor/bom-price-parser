# Octopart BOM Tool

- URL: `https://octopart.com/bom-tool/`
- Требуется аккаунт Nexar/Octopart. Если не залогинен — остановиться и
  попросить пользователя войти.

## Процедура (уточнить при первом прогоне)

1. Открыть URL, подождать 4–6 с.
2. Создать новый BOM, загрузить CSV (`file_upload`), сопоставить колонки
   MPN / Qty / Manufacturer.
3. В настройках отображения выбрать **авторизованных дистрибьюторов**
   (Authorized only, если есть такой фильтр) и валюту **USD**.
4. Дождаться подбора, экспортировать в CSV/xlsx.

## Формат экспорта

Не верифицирован (заглушка в `SITES`). Octopart обычно отдаёт отдельную
колонку подобранного MPN (`Manufacturer Part Number`) — тогда merge работает
без `--trust-echo`. Если merge упадёт с BrowserExportFormatError — добавить
недостающие синонимы в `SITES['octopart']['aliases']` в `browser_bom.py`
(текст ошибки подскажет какие) и положить копию экспорта в
`final/browser/samples/`.

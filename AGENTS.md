# bom-price-parser — цены на ЭКБ по BOM

## Purpose

Инструмент подбирает цену по каждой позиции BOM-запроса и заполняет итоговый xlsx по
эталону. Основной способ — официальные API дистрибьюторов (DigiKey, Mouser, TME,
Newark/Farnell). Для каждой позиции выбирается **одно** предложение.

Это не обзор рынка: если нужны все предложения по каждой позиции сразу и date code —
задача соседнего проекта `Project C\count pars`, а не этого.

## Routing

| Сигналы в запросе | Читать |
|---|---|
| прогнать BOM, посчитать цены, обычный запуск, флаги, одна позиция | [`.agents/workflows/api-search.md`](.agents/workflows/api-search.md) |
| API не дал результата, добрать вручную, oemsecrets, Cloudflare | [`.agents/workflows/browser-fallback.md`](.agents/workflows/browser-fallback.md) |
| можно ли взять этот артикул, аналог, суффикс упаковки, маска `x`, производитель | [`.agents/rules/part-matching.md`](.agents/rules/part-matching.md) |
| какое предложение выбрать, MOQ, брекет, сток, Lead Time, когда это RFQ | [`.agents/rules/price-selection.md`](.agents/rules/price-selection.md) |
| колонки, стили, заливки, формат цены, сноска, рубли | [`.agents/rules/output-format.md`](.agents/rules/output-format.md) |
| ключи, домены, зависимости | [`.agents/references/environment.md`](.agents/references/environment.md) |

Сквозной прогон: workflow → part-matching → price-selection → output-format.

## Priority

При совпадении нескольких маршрутов побеждает верхний.
API-маршрут основной; браузерный подключается только к позициям, которые API не закрыл.
Явная инструкция пользователя перекрывает умолчания, но не Global constraints.

## Fallback

Ни один маршрут не подошёл — прочитать
[`.agents/workflows/api-search.md`](.agents/workflows/api-search.md) и уточнить у
пользователя. Не угадывать производителя и не подбирать аналог.

## Loading rule

Read only the Markdown files required by the matched route.

## Global constraints

- `--yes` обязателен при запуске из-под агента: без него скрипт встаёт на `input()`.
- Аналог с другими характеристиками недопустим. Отличаться может только суффикс упаковки.
- Производитель не указан либо не совпал — это `RFQ`, а не подбор замены.
- Соседний проект `Project C\count pars` не изменять.
- `update_bom*.py` — архив прошлых заявок с зашитыми данными; для новых задач не применять.

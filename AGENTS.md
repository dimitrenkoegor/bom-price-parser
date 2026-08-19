# BOM Price Parser — цены на ЭКБ по BOM

## Purpose

Автопоиск цен на электронные компоненты: письмо с BOM → парсер → API дистрибьюторов →
ответ письмом. Для каждой позиции выбирается **одно** предложение. Что не закрыли API,
добирается браузерным каскадом по BOM-инструментам площадок.

Это не обзор рынка: если нужны все предложения по каждой позиции сразу и date code —
задача соседнего проекта `..\count pars`, а не этого.

## Routing

| Сигналы в запросе | Читать |
|---|---|
| прогнать BOM, посчитать цены, обычный запуск, флаги, одна позиция | [`.agents/workflows/api-search.md`](.agents/workflows/api-search.md) |
| добить RFQ-остаток, каскад, oemsecrets, octopart, findchips, trustedparts | [`.agents/workflows/browser-cascade.md`](.agents/workflows/browser-cascade.md) |
| почта, автоответ, монитор не отвечает, `mail-state.json`, автозапуск | [`.agents/workflows/mail-monitor.md`](.agents/workflows/mail-monitor.md) |
| можно ли взять этот артикул, аналог, суффикс упаковки, маска `x`, производитель | [`.agents/rules/part-matching.md`](.agents/rules/part-matching.md) |
| какое предложение выбрать, MOQ, брекет, сток, Lead Time, когда это RFQ | [`.agents/rules/price-selection.md`](.agents/rules/price-selection.md) |
| колонки, стили, что значит персиковая и сиреневая строка, рубли | [`.agents/rules/output-format.md`](.agents/rules/output-format.md) |
| ключи, флаги `.env`, домены, чего ещё ждём | [`.agents/references/environment.md`](.agents/references/environment.md) |
| пошаговая браузерная нога по конкретной площадке | `.claude/skills/bom-browser-cascade/SKILL.md` |

Сквозной прогон: api-search → part-matching → price-selection → output-format,
затем при остатке RFQ — browser-cascade.

## Priority

При совпадении нескольких маршрутов побеждает верхний.
API-маршрут основной; каскад подключается только к тому, что API не закрыл.
Явная инструкция пользователя перекрывает умолчания, но не Global constraints.

## Fallback

Ни один маршрут не подошёл — прочитать
[`.agents/workflows/api-search.md`](.agents/workflows/api-search.md) и уточнить у
пользователя. Не угадывать производителя и не подбирать аналог.

## Loading rule

Read only the Markdown files required by the matched route.

## Global constraints

- `--yes` обязателен при запуске из-под агента: без него скрипт встаёт на `input()`.
- Итоговый xlsx и CSV каскада руками не править: цена входит только через `merge`.
- Аналог с другими характеристиками недопустим. Отличаться может только суффикс упаковки.
- Производитель не указан либо не совпал — это `RFQ`, а не подбор замены.
- Монитор почты запускать **одним** способом; `mail-state.json` не удалять.
- Соседний проект `..\count pars` не изменять.

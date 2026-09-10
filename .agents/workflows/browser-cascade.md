# Запасной маршрут: браузерный каскад по BOM-инструментам

Добирает RFQ-остаток после того, как отработал API-маршрут
([`api-search.md`](api-search.md)). Каскадом: каждая следующая площадка получает только то,
что не закрыла предыдущая. Порядок — `oemsecrets` → `octopart` → `findchips` → `trustedparts`.

## Кто чем владеет

- **CLI** — `browser_bom.py` с тремя командами: `export`, `merge`, `status`.
- **Процедура браузерной ноги** — скилл `.claude/skills/bom-browser-cascade/SKILL.md`
  и его `references/<site>.md` по каждой площадке. Здесь их копий нет.

Правила выбора цены и сопоставления артикула те же, что и на основном маршруте:
[`../rules/price-selection.md`](../rules/price-selection.md),
[`../rules/part-matching.md`](../rules/part-matching.md).
Строки, добытые каскадом, красятся сиреневым —
[`../rules/output-format.md`](../rules/output-format.md).

## Цикл на одну площадку

```bash
python browser_bom.py export --site <site> --result "<итоговый.xlsx>"
python browser_bom.py merge  --site <site> --result "<итоговый.xlsx>" --export "<скачанный>" --dry-run
python browser_bom.py status --result "<итоговый.xlsx>"
```

Полный порядок шагов, разбор вывода и жёсткие правила прохода — в SKILL.md каскада.

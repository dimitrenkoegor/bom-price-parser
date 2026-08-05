# oemsecrets BOM Tool

- URL: `https://bomtool.oemsecrets.com/`
- Требуется аккаунт (бесплатный). Если не залогинен — остановиться и
  попросить пользователя войти.

## Процедура

1. Открыть URL, подождать 4–6 с (данные грузятся асинхронно).
2. Создать новый BOM / «Upload BOM», загрузить CSV через
   `mcp__claude-in-chrome__file_upload`.
3. В диалоге сопоставления колонок: `Part Number` -> Part Number,
   `Quantity` -> Quantity, `Manufacturer` -> Manufacturer.
4. Открыть **Distributor Preferences** и оставить только авторизованных:
   DigiKey, Mouser, Arrow, TTI, TME, Avnet, Newark/Farnell, Future, Heilind,
   RS Components, Rutronik. Verical допустим (= Arrow).
5. Выставить валюту **USD** (сайт по умолчанию показывает EUR; экспортная
   колонка называется `Unit Price in USD` — проверить, что это так).
6. Запустить **Data Autofill**, дождаться окончания прогресс-бара
   (не трогать вкладку, может занять минуты).
7. Export -> xlsx. Запомнить путь скачанного файла.

## Формат экспорта (верифицирован)

9 колонок: `Part Number`, `Quantity for Single BOM`, `Manufacturer`,
`Distributor`, `Minimum Order`, `Stock`,
`Lead Time on Additional Stock in Weeks`, `Unit Price in USD`,
`Distributor SKU`.

Колонки подобранного MPN нет — экспорт **echo-only**, merge запускать
с `--trust-echo` (производитель в экспорте есть и валидируется).

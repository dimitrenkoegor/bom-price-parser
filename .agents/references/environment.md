# Окружение, ключи, открытые задачи

## Зависимости

Python 3.x и `openpyxl`:

```bash
pip install openpyxl
```

## Ключи API

Файл `.env` рядом со скриптом. В git он не хранится и лежит только на рабочей машине:

```text
MOUSER_API_KEY
DIGIKEY_CLIENT_ID
DIGIKEY_CLIENT_SECRET
TME_APP_SECRET
TME_TOKEN
FARNELL_API_KEY
OEMSECRETS_API_KEY   # добивочный фолбэк, 10 позиций/день
```

Без ключей скрипт не падает, но все позиции уходят в `RFQ`.

## Флаги поведения

- `HOLD_ON_RFQ` — придержать почтовый ответ при остатке RFQ, см.
  [`../workflows/mail-monitor.md`](../workflows/mail-monitor.md).

## Нужен доступ в интернет

- `api.digikey.com`
- `api.mouser.com`
- `api.tme.eu`
- `api.element14.com`
- `www.cbr.ru` — курс рубля
- `open.er-api.com` — запасной источник курса

## Открытые задачи

- Ждём ключ **Arrow API** (заявка от 16.07) — станет пятым источником, без дневных квот.
- Нужен ключ **oemsecrets** (`OEMSECRETS_API_KEY`) — добивочный фолбэк, 10 позиций в сутки.

## Соседние проекты

- `..\count pars` — обзор рынка: показывает **все** предложения по позиции для сравнения,
  а не выбирает одно лучшее. Этот проект он читает как справочник рабочих форм запросов.
- `C:\Users\Huawei\Desktop\Project\bom-price-parser` — второй рабочий чекаут того же
  репозитория, отставший по коду. Живой — этот.

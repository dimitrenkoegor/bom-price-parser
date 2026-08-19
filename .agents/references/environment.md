# Окружение и ключи

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
```

Без ключей скрипт не падает, но все позиции уходят в `RFQ`.

## Нужен доступ в интернет

- `api.digikey.com`
- `api.mouser.com`
- `api.tme.eu`
- `api.element14.com`
- `www.cbr.ru` — курс рубля
- `open.er-api.com` — запасной источник курса

## Соседний проект

`Project C\count pars` решает другую задачу: показывает **все** предложения по позиции для
сравнения, а не выбирает одно лучшее. У него свой `.env`. Этот проект для него — только
справочник рабочих форм запросов, читается, но не изменяется.

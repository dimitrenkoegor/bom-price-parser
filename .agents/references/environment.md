# Окружение, ключи, структура

## Зависимости

Python 3.11+ (stdlib `tomllib`) и `openpyxl`:

```bash
pip install openpyxl
```

Всё остальное — стандартная библиотека (`urllib`, `imaplib`, `smtplib`). Браузерных
библиотек (selenium, playwright, bs4) в проекте нет и быть не должно.

## Структура

```
price_parser.py          точка входа (флаги — .agents/workflows/api-search.md)
mail_monitor.py          почтовый демон (.agents/workflows/mail-monitor.md)
src/bomprice/
  store.py  net.py  fx.py            пути, .env, TOML, кэш, квоты; HTTP; курсы
  mpn.py                             разбор артикула и качество совпадения — ОБЩИЙ с count pars
  offer.py  registry.py              запись предложения; реестр площадок из sources.toml
  adapters/api_franchise.py          mouser, digikey, farnell (newark), tme
  adapters/api_aggregator.py         oemsecrets, nexar, arrow
  bomread.py                         Excel/txt → позиции, превью и обратно
  collect.py                         опрос площадок по позиции + второй проход точными артикулами
  select.py                          регламент отбора одного предложения
  report.py                          10-колоночный xlsx + сноска
  llm.py                             LLM-контур (headless claude) с проверкой через API
  cli.py                             разбор флагов и поток прогона
config/settings.toml     темп, кэш, курс, [selection]
config/sources.toml      площадки: только tier = "api"
config/brands.toml       словарь брендов — ОБЩИЙ с count pars
tests/                   unittest; test_mpn.py общий с count pars
archive/                 старые заявки с зашитыми данными — не использовать
work/                    кэш ответов, квоты, курсы (удаляемо)
start/  final/           вход / выход
```

Общие с count pars файлы (`mpn.py`, `brands.toml`, `test_mpn.py`, адаптеры API) должны
совпадать байт в байт: правка в одном проекте копируется во второй.

## Ключи и настройки `.env`

Файл `.env` в корне, в репозиторий не попадает; полный список с комментариями —
`.env.example`. Группы:

| Группа | Переменные |
|---|---|
| API основного прохода | `DIGIKEY_CLIENT_ID`, `DIGIKEY_CLIENT_SECRET`, `MOUSER_API_KEY`, `FARNELL_API_KEY`, `TME_TOKEN`, `TME_APP_SECRET`, `TME_CUSTOMER_NUMBER` |
| добивка RFQ | `OEMSECRETS_API_KEY`, `OEMSECRETS_DAILY_LIMIT`, `NEXAR_CLIENT_ID`, `NEXAR_CLIENT_SECRET`, `ARROW_API_KEY`, `ARROW_LOGIN` |
| LLM-контур | `ENABLE_CLAUDE_FALLBACK`, `CLAUDE_MAX_ITEMS`, `CLAUDE_BIN` |
| почта | `MAIL_IMAP_HOST`, `MAIL_IMAP_MAILBOX`, `MAIL_USER`, `MAIL_PASSWORD`, `MAIL_SMTP_HOST`, `MAIL_SMTP_PORT`, `ALLOWED_SENDERS`, `POLL_SECONDS`, `SEND_ENABLED`, `HOLD_ON_RFQ` |

Без ключей скрипт не падает, но все позиции уходят в `RFQ`. Готовность площадок одним
экраном — `ПРОВЕРКА.bat`.

## Нужен доступ в интернет

- `api.digikey.com`, `api.mouser.com`, `api.tme.eu`, `api.element14.com`
- `oemsecretsapi.com` (добивка), `api.nexar.com`, `identity.nexar.com` (если включён)
- `www.cbr.ru` — курс рубля; `open.er-api.com` — кросс-курсы
- `imap.mail.ru`, `smtp.mail.ru` — почтовый монитор

## Диагностика

| Симптом | Причина | Что делать |
|---|---|---|
| окно bat пишет `'жет' is not recognized ...` — обрывки русских слов | старые bat с кириллицей и `chcp 65001`: cmd читает их с середины строк | скачать репозиторий заново; в `.bat` кириллицу не писать, текст — в `_launcher.ps1` (UTF-8 с BOM) |
| `нет ключа` в `ПРОВЕРКА.bat` | пустая переменная | заполнить `.env` по `.env.example` |
| `HTTP 401` | ключ протух или неверен | перевыпустить в кабинете площадки |
| `HTTP 403` у Mouser | слишком часто | темп уже 0.5 rps; подождать, снизить `rate.rps` в `config/sources.toml` |
| `HTTP 429` | превышен темп | снизить `rate.rps` |
| `суточная квота исчерпана` | `rate.daily` / `OEMSECRETS_DAILY_LIMIT` | подождать суток; счётчик — `work/quota.json` |
| все позиции RFQ «производитель не указан» | не распозналась колонка производителя | открыть превью, вписать бренд или подсказать `--mfr-col` |
| цена есть, а нужна свежая | ответ взят из кэша | `--no-cache` либо снести `work/cache` |
| `в книге нет листа …` | опечатка в `--sheet` | взять имя из списка в сообщении |
| колонки RUB пустые | ЦБ не ответил | `--rub-rate 85` |

Сырые ответы площадок — `work/cache/<площадка>/`; в контекст модели их целиком не тащить.

## Соседний проект

`C:\Users\Huawei\Desktop\Project C\count pars` — обзор рынка: **все** предложения по позиции
со всего спектра площадок (API, HTTP-парсеры, браузер, выгрузки), с брокерами и date code.
У него свой `.env` и свой `work/`. Этот проект для него — справочник форм запросов;
изменяется только копированием общих файлов.

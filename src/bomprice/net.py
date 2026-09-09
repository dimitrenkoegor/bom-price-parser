"""HTTP поверх stdlib urllib: единый User-Agent, ретраи, троттлинг по хосту.

Зависимостей нет намеренно — проект должен запускаться на голом Python.
"""

from __future__ import annotations

import gzip
import json
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import zlib

from . import store

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

DEFAULT_HEADERS = {
    "User-Agent": UA,
    "Accept-Language": "en-US,en;q=0.9,ru;q=0.8",
    "Accept-Encoding": "gzip, deflate",
}


class HttpError(Exception):
    def __init__(self, status: int, message: str = ""):
        super().__init__(f"HTTP {status}: {message}"[:300])
        self.status = status


class _HostThrottle:
    """Не больше rps запросов к хосту и не больше одного запроса в полёте."""

    def __init__(self) -> None:
        self._locks: dict[str, threading.Lock] = {}
        self._next_at: dict[str, float] = {}
        self._guard = threading.Lock()

    def _lock_for(self, host: str) -> threading.Lock:
        with self._guard:
            return self._locks.setdefault(host, threading.Lock())

    def wait(self, host: str, rps: float):
        lock = self._lock_for(host)
        lock.acquire()
        gap = 1.0 / rps if rps and rps > 0 else 0.0
        now = time.monotonic()
        due = self._next_at.get(host, 0.0)
        if due > now:
            time.sleep(due - now)
        self._next_at[host] = max(now, due) + gap
        return lock


_throttle = _HostThrottle()


def _decode(resp) -> bytes:
    raw = resp.read()
    enc = (resp.headers.get("Content-Encoding") or "").lower()
    if enc == "gzip":
        try:
            return gzip.decompress(raw)
        except OSError:
            return raw
    if enc == "deflate":
        try:
            return zlib.decompress(raw)
        except zlib.error:
            try:
                return zlib.decompress(raw, -zlib.MAX_WBITS)
            except zlib.error:
                return raw
    return raw


def fetch(url: str, *, data: bytes | None = None, headers: dict | None = None,
          method: str | None = None, rps: float = 1.0, timeout: int | None = None,
          retries: int | None = None, context=None) -> bytes:
    """Один HTTP-запрос. Ретраит только 429/5xx и сетевые сбои — 4xx поднимаем сразу.

    context — свой ssl.SSLContext. Нужен российским площадкам: они перешли на
    отечественные сертификаты, которых нет в системном хранилище Windows.
    """
    cfg = store.settings().get("run", {})
    timeout = timeout or cfg.get("timeout", 25)
    retries = cfg.get("retries", 2) if retries is None else retries

    hdrs = dict(DEFAULT_HEADERS)
    if headers:
        hdrs.update(headers)
    host = urllib.parse.urlsplit(url).netloc
    last: Exception | None = None

    for attempt in range(1, retries + 2):
        lock = _throttle.wait(host, rps)
        try:
            req = urllib.request.Request(url, data=data, headers=hdrs, method=method)
            with urllib.request.urlopen(req, timeout=timeout, context=context) as resp:
                return _decode(resp)
        except urllib.error.HTTPError as exc:
            body = ""
            try:
                body = exc.read()[:300].decode("utf-8", "replace")
            except Exception:
                pass
            if exc.code in (429,) or 500 <= exc.code < 600:
                last = HttpError(exc.code, body)
            else:
                raise HttpError(exc.code, body) from None
        except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as exc:
            last = exc
        finally:
            lock.release()
        if attempt <= retries:
            time.sleep(attempt * 1.5)

    raise last if last else HttpError(0, "неизвестная сетевая ошибка")


def get_json(url: str, **kw):
    kw.setdefault("headers", {}).setdefault("Accept", "application/json")
    return json.loads(fetch(url, **kw).decode("utf-8", "replace"))


def post_json(url: str, payload, **kw):
    headers = kw.pop("headers", {})
    headers.setdefault("Content-Type", "application/json")
    headers.setdefault("Accept", "application/json")
    body = json.dumps(payload).encode("utf-8")
    return json.loads(fetch(url, data=body, headers=headers, method="POST", **kw)
                      .decode("utf-8", "replace"))


def post_form(url: str, fields: dict, **kw):
    headers = kw.pop("headers", {})
    headers.setdefault("Content-Type", "application/x-www-form-urlencoded")
    body = urllib.parse.urlencode(fields).encode("utf-8")
    return fetch(url, data=body, headers=headers, method="POST", **kw)


def get_text(url: str, **kw) -> str:
    return fetch(url, **kw).decode("utf-8", "replace")

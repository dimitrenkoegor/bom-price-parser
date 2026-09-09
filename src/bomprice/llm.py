"""LLM-контур: подсказка написания артикула + ОБЯЗАТЕЛЬНАЯ проверка через API.

Включается ENABLE_CLAUDE_FALLBACK=true в .env. Headless Claude CLI получает
RFQ-остаток и возвращает JSON с предполагаемым полным артикулом и брендом; цены
у него не спрашиваются. В файл попадает только то, что подтвердилось ценой у
дистрибьютора по обычному регламенту (select.best); такие строки помечаются
llm_assisted и красятся персиковым, чтобы закупщик их перепроверил.
"""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile

from . import bomread, collect as collectmod, select as selectmod, store

SYSTEM = (
    "Ты — детерминированный резолвер артикулов ЭКБ. Отвечай ТОЛЬКО валидным "
    "JSON-массивом в одну выдачу, без приветствий, пояснений и markdown. "
    "Никогда не задавай уточняющих вопросов. Идентифицируй ТОТ ЖЕ компонент, "
    "не аналог: другой номинал/напряжение/корпус/точность запрещены. Разрешено "
    "исправить опечатку, дописать префикс/суффикс серии, указать актуальный "
    "бренд после поглощений, развернуть неполный артикул. Цены/склад/сроки не "
    "придумывай. Не уверен — suggested_mpn=null."
)

PROMPT = """Для каждой позиции верни объект. Формат ответа — только этот JSON-массив:
[{"requested_pn":"<как в запросе>","suggested_mpn":"<полный артикул>"|null,"suggested_manufacturer":"<бренд>"|null,"reason":"<кратко>"}]

Позиции:
"""


def binary() -> str:
    """Путь к CLI. На Windows npm ставит обёртку claude.cmd, которую subprocess
    без расширения не находит."""
    configured = store.env("CLAUDE_BIN")
    if configured:
        return configured
    for name in ("claude.cmd", "claude.exe", "claude"):
        found = shutil.which(name)
        if found:
            return found
    return "claude"


def suggest(items: list[dict], timeout: int = 180, log=print) -> list[dict]:
    """Спрашивает headless-Claude про нерешённые позиции. → список подсказок (может быть пустым)."""
    payload = [{"requested_pn": it["pn"], "manufacturer": it.get("manufacturer", ""),
                "description": (it.get("description") or "")[:120]} for it in items]
    prompt = PROMPT + json.dumps(payload, ensure_ascii=False, indent=1)
    # Пустая временная папка: CLI не подхватит AGENTS.md/CLAUDE.md проекта и не
    # «уйдёт в диалог». Промпт — через stdin: .cmd-обёртка ломает кавычки в аргументах.
    try:
        with tempfile.TemporaryDirectory() as workdir:
            completed = subprocess.run(
                [binary(), "-p", "--output-format", "json", "--append-system-prompt", SYSTEM],
                input=prompt, capture_output=True, text=True, timeout=timeout,
                encoding="utf-8", errors="replace", cwd=workdir)
    except (OSError, subprocess.TimeoutExpired) as exc:
        log(f"    LLM-контур недоступен: {exc}")
        return []
    raw = completed.stdout or ""
    if not raw.strip():
        log(f"    LLM-контур: пустой ответ ({(completed.stderr or '').strip()[:80]})")
        return []
    text = raw
    try:
        wrapper = json.loads(raw)
        if isinstance(wrapper, dict):
            text = wrapper.get("result") or wrapper.get("text") or raw
    except json.JSONDecodeError:
        pass
    match = re.search(r"\[.*\]", text, re.S)
    if not match:
        return []
    try:
        data = json.loads(match.group(0))
    except json.JSONDecodeError:
        return []
    return data if isinstance(data, list) else []


def assisted_pass(results: list[dict], parts, sources, sources_by_id, rates, stats,
                  use_cache: bool = True, log=print) -> int:
    """Для RFQ-остатка: подсказка LLM → проверка через API → замена строки. → сколько добрано."""
    if not store.env_flag("ENABLE_CLAUDE_FALLBACK"):
        return 0
    pending = [(i, r) for i, r in enumerate(results) if r.get("status") == "RFQ" and r.get("pn")]
    if not pending:
        return 0
    limit = int(store.env("CLAUDE_MAX_ITEMS", "40") or 40)
    batch = [r for _, r in pending[:limit]]
    log(f"\nLLM-контур: уточняю {len(batch)} нерешённых позиций...")
    suggestions = suggest(batch, log=log)
    if not suggestions:
        log("LLM-контур: подсказок нет")
        return 0

    from . import mpn as mpnmod
    by_pn = {}
    for s in suggestions:
        if isinstance(s, dict) and s.get("requested_pn"):
            by_pn[mpnmod.canonical(s["requested_pn"])] = s

    accept = selectmod.accept_set()
    recovered = 0
    for i, r in pending:
        s = by_pn.get(mpnmod.canonical(r["pn"]))
        if not s:
            continue
        mpn = (s.get("suggested_mpn") or "").strip()
        manu = (s.get("suggested_manufacturer") or r.get("manufacturer") or "").strip()
        if not mpn or not manu:
            continue                      # без бренда цену не подставляем (регламент)
        probe = bomread.literal_part(parts[i].part_id, mpn, r.get("qty", 1), manu,
                                     r.get("description", ""))
        offers = collectmod.collect_with_resolution(probe, sources, stats, accept, use_cache, log)
        cand = selectmod.best(offers, probe, rates, sources_by_id)
        if cand.get("status") != "FOUND":
            continue
        cand["pn"] = r["pn"]                      # в колонке A остаётся запрошенное написание
        cand["resolved_mpn"] = cand.get("resolved_mpn") or mpn
        cand["description"] = r.get("description", r["pn"])
        cand["llm_assisted"] = True
        cand["llm_reason"] = (s.get("reason") or "")[:120]
        results[i] = cand
        recovered += 1
        log(f"    LLM+API: {r['pn']} -> {cand['resolved_mpn']} | {cand['distributor']} "
            f"${cand['price_usd']}  ({cand['llm_reason']})")
    log(f"LLM-контур подтвердил позиций: {recovered} из {len(batch)}")
    return recovered

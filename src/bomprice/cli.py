"""CLI котировщика. Флаги совместимы с прежним price_parser.py — их вызывают
bat-файлы и mail_monitor.py (--input --output --preview-path --yes).

Поток:
  1. позиции из Excel/txt (или одна из --once) → превью → пауза без --yes;
  2. курс ЦБ; готовность площадок по .env;
  3. по каждой позиции: опрос франшизных API → регламент отбора (select);
  4. RFQ-остаток: агрегаторы (oemsecrets, в пределах квоты), затем LLM-контур;
  5. xlsx по эталону + две строки итога для монитора.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import re
import sys
from pathlib import Path

from . import bomread, collect as collectmod, fx as fxmod, llm, registry, report, select as selectmod, store


def _col(text: str) -> int:
    text = (text or "").strip().upper()
    if not text or text == "0":
        return 0
    if text.isdigit():
        return int(text)
    if re.fullmatch(r"[A-Z]{1,3}", text):
        num = 0
        for ch in text:
            num = num * 26 + (ord(ch) - ord("A") + 1)
        return num
    raise argparse.ArgumentTypeError(f"колонка задаётся буквой (B) или числом (2), не {text!r}")


def load_items(args, log=print) -> list:
    if args.once:
        return [bomread.literal_part(1, args.once[0], int(args.once[1]), args.manufacturer or "")]
    if args.input:
        path = Path(args.input)
        if not path.is_absolute():
            path = store.ROOT / path
    else:
        path = bomread.autopick_request()
        if path is None:
            txt = store.ROOT / "input.txt"
            if txt.exists():
                log(f"Excel-запрос не найден, читаю {txt.name}")
                return bomread.read_txt(txt)
            raise SystemExit("Не найден ни Excel-запрос, ни input.txt. Положите файл запроса в start/.")
        log(f"Авто-выбран запрос: {path.name}")
    if not path.exists():
        raise SystemExit(f"Файл не найден: {path}")
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        return bomread.read_bom(path, sheet=args.sheet, header_row=args.header_row,
                                pn_col=args.pn_col, name_col=args.name_col,
                                qty_col=args.qty_col, mfr_col=args.mfr_col, log=log)
    if path.suffix.lower() == ".xls":
        raise SystemExit(f"{path.name}: формат .xls не читается — пересохраните как .xlsx")
    return bomread.read_txt(path)


def print_preview(parts, preview_path, log=print) -> None:
    log(f"\nРаспознано позиций: {len(parts)}. Превью: {preview_path}")
    for part in parts[:30]:
        flag = ""
        if not part.mpn:
            flag = "  <-- ПУСТО, впишите артикул!"
        elif not part.manufacturer:
            flag = "  <-- нет производителя: уйдёт в RFQ"
        log(f"  #{part.part_id:>3}  {part.mpn or '(нет)':<28} x{part.qty:<6} {part.manufacturer}{flag}")
    if len(parts) > 30:
        log(f"  ... ещё {len(parts) - 30}")


def run(args, log=print) -> int:
    preview_path = Path(args.preview_path) if args.preview_path else store.PREVIEW

    # 1) позиции
    if args.from_preview and preview_path.exists():
        parts = bomread.read_preview(preview_path)
        log(f"Читаю из превью: {preview_path.name}")
    else:
        parts = load_items(args, log)
    if args.limit:
        parts = parts[: args.limit]

    # 2) превью + чекпоинт
    if not args.once and not args.from_preview:
        bomread.write_preview(parts, preview_path)
        print_preview(parts, preview_path, log)
        if args.preview_only:
            log("\nРежим превью: файл сохранён, цены не искались.")
            return 0
        if not args.yes:
            log(f"\nПроверьте/поправьте {preview_path.name} в Excel и СОХРАНИТЕ его.")
            answer = input("Нажмите Enter для поиска цен, или N для выхода: ").strip().lower()
            if answer == "n":
                log("Остановлено. Поправьте превью и запустите снова.")
                return 0
            if preview_path.exists():
                parts = bomread.read_preview(preview_path)

    # 3) курс и площадки
    rates = fxmod.load(rub_rate=args.rub_rate or 0.0, markup=args.rub_markup)
    log(f"USD/RUB: {rates.usd_rub:.4f} ({rates.source}, {rates.at})")
    main_sources = collectmod.usable_sources(("franchise",))
    fallback_sources = collectmod.usable_sources(("aggregator",))
    all_sources = registry.load_sources()
    sources_by_id = {s.id: s for s in all_sources}
    if not main_sources:
        log("\nВНИМАНИЕ: ни одна площадка не готова (нет ключей в .env?). Все позиции уйдут в RFQ.")
        for src in all_sources:
            state, why = registry.readiness(src)
            log(f"  {src.id:<12} {state} {why}")
    else:
        log(f"\nАктивные источники: {', '.join(s.title for s in main_sources)}"
            + (f"; добивка RFQ: {', '.join(s.title for s in fallback_sources)}" if fallback_sources else ""))

    # 4) поиск
    accept = selectmod.accept_set()
    stats = collectmod.Stats()
    use_cache = not args.no_cache
    results: list[dict] = [None] * len(parts)  # type: ignore[list-item]
    log(f"Позиций к поиску: {len(parts)}")

    def work(index: int) -> tuple[int, dict, list[str]]:
        part = parts[index]
        lines = [f"[{index + 1}/{len(parts)}] #{part.part_id}: {part.mpn or '(нет артикула)'}  qty={part.qty}"]
        if not part.mpn:
            return index, selectmod.rfq(part, "артикул не распознан"), lines + ["    -> RFQ: артикул не распознан"]
        if not part.manufacturer:
            return index, selectmod.rfq(part, "производитель не указан"), lines + ["    -> RFQ: производитель не указан"]
        offers = collectmod.collect_with_resolution(part, main_sources, stats, accept, use_cache,
                                                    log=lines.append)
        result = selectmod.best(offers, part, rates, sources_by_id)
        if result["status"] == "FOUND":
            tag = "В наличии" if result["in_stock"] else "под заказ"
            lines.append(f"    -> {result['distributor']}  ${result['price_usd']:.3f}  "
                         f"({tag}, склад {result['stock']}, MOQ {result['moq']})")
        else:
            lines.append(f"    -> RFQ ({result.get('reason', '')[:100]})")
        return index, result, lines

    workers = max(1, int(store.settings().get("run", {}).get("workers", 4)))
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        for index, result, lines in pool.map(work, range(len(parts))):
            results[index] = result
            log("\n".join(lines))

    # 5) добивка RFQ: агрегаторы в пределах квоты, позиции с бо́льшим количеством первыми
    if fallback_sources:
        pending = sorted([i for i, r in enumerate(results) if r["status"] == "RFQ"
                          and parts[i].mpn and parts[i].manufacturer],
                         key=lambda i: -parts[i].qty)
        recovered = 0
        for i in pending:
            offers = collectmod.collect_part(parts[i], fallback_sources, stats, use_cache, log)
            cand = selectmod.best(offers, parts[i], rates, sources_by_id)
            if cand["status"] == "FOUND":
                results[i] = cand
                recovered += 1
                log(f"    {cand['source_id']}: {parts[i].mpn} -> {cand['distributor']} ${cand['price_usd']}")
        if recovered:
            log(f"Добивка через агрегаторы: найдено ещё {recovered}")

    # 6) LLM-контур
    try:
        llm.assisted_pass(results, parts, main_sources, sources_by_id, rates, stats, use_cache, log)
    except Exception as exc:                          # noqa: BLE001
        log(f"LLM-контур пропущен: {exc}")

    log("\n" + collectmod.summary(stats))
    out_path = Path(args.output) if args.output else store.DEFAULT_OUTPUT
    report.write_results(results, out_path, rates, [s.title for s in main_sources] or ["—"], log)
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="price_parser.py",
        description="Цены на ЭКБ по BOM через API дистрибьюторов (DigiKey, Mouser, TME, Newark/Farnell; "
                    "oemsecrets для RFQ-остатка)")
    ap.add_argument("--input", "-i", default=None, help="Excel/txt запрос. По умолч. авто-выбор из start/")
    ap.add_argument("--output", "-o", default=None, help="итоговый xlsx (по умолч. final/BOM_Приложение_1.xlsx)")
    ap.add_argument("--sheet", default=None, help="имя листа Excel; нет — первый")
    ap.add_argument("--header-row", type=int, default=0, help="строка шапки, если не нашлась сама")
    ap.add_argument("--pn-col", type=_col, default=0, help="колонка артикула: B или 2")
    ap.add_argument("--name-col", type=_col, default=0, help="колонка наименования")
    ap.add_argument("--qty-col", type=_col, default=0, help="колонка количества")
    ap.add_argument("--mfr-col", type=_col, default=0, help="колонка производителя")
    ap.add_argument("--once", nargs=2, metavar=("PN", "QTY"), help="одна позиция для проверки")
    ap.add_argument("--manufacturer", default="", help="производитель для режима --once")
    ap.add_argument("--rub-rate", type=float, default=None, help="курс USD/RUB вручную (иначе ЦБ)")
    ap.add_argument("--rub-markup", type=float, default=None,
                    help="наценка к курсу USD/RUB (по умолч. из config/settings.toml, +4)")
    ap.add_argument("--preview-only", action="store_true", help="только распознать в превью, без цен")
    ap.add_argument("--from-preview", action="store_true", help="сразу читать превью_позиции.xlsx")
    ap.add_argument("--preview-path", type=Path, default=None,
                    help="путь к превью; нужен для изолированной обработки очереди")
    ap.add_argument("--yes", "-y", action="store_true", help="не спрашивать подтверждение")
    ap.add_argument("--limit", type=int, default=0, help="взять только первые N позиций")
    ap.add_argument("--no-cache", action="store_true", help="не брать ответы площадок из work/cache")
    args = ap.parse_args(argv)
    try:
        return run(args)
    except KeyboardInterrupt:
        print("\nПрервано.")
        return 130


if __name__ == "__main__":
    sys.exit(main())

"""bom-price-parser — цены на ЭКБ по BOM через официальные API дистрибьюторов.

Точка входа. Весь код — в src/bomprice/ (общее ядро с count pars + регламент
отбора одного предложения). Флаги: python price_parser.py --help.

    python price_parser.py --input "start\\запрос.xlsx" --output final\\result.xlsx --yes
    python price_parser.py --once C5750X7R1H106KT 320 --manufacturer TDK --yes
    python price_parser.py --preview-only
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

if sys.platform == "win32":
    # консоль Windows по умолчанию cp866/cp1251 — русский вывод и итог для монитора
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

from bomprice.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())

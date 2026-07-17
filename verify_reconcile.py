"""Проверка логики сверки без графического окна."""

from __future__ import annotations

import sys
from pathlib import Path


def main() -> int:
    from parser_card import parse_card
    from parser_cash import parse_cash
    from parser_ofd import combine_ofd_sources, parse_ofd_checks, parse_ofd_shifts
    from reconcile import reconcile
    from excel_report import generate_report

    downloads = Path.home() / "Downloads"
    ofd1 = downloads / "Отчет по чекам ИНН 5911054930 2026-06-01 - 2026-06-30 (1).xlsx"
    ofd2 = downloads / "Отчет по чекам ИНН 5911054930 2026-06-01 - 2026-06-30 (2).xlsx"
    card = downloads / "реестр оплат платежной картой_июнь26.xlsx"
    cash = downloads / "реестр кассовых документов_06,26_Березники.xlsx"

    missing = [str(p) for p in (ofd1, card, cash) if not p.exists()]
    if missing:
        print("Пропуск: тестовые файлы не найдены в Downloads.")
        return 0

    _, main_ofd = parse_ofd_checks(str(ofd1))
    shift = parse_ofd_shifts(str(ofd2)) if ofd2.exists() else None
    ofd = combine_ofd_sources(main_ofd, shift) if shift is not None else main_ofd
    _, card_g = parse_card(str(card))
    _, cash_g = parse_cash(str(cash))

    result = reconcile(ofd, card_g, cash_g)
    out = downloads / "_test_sverka_output.xlsx"
    generate_report(result, str(out))

    assert out.exists(), "Отчёт не создан"
    assert len(result.daily) > 0, "Нет строк сверки"
    out.unlink(missing_ok=True)

    print("OK: сверка и сохранение Excel работают.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""Парсер отчёта по чекам ОФД."""

from __future__ import annotations

import pandas as pd

from config import (
    COLUMN_MAPPINGS,
    EXCLUDED_DEPARTMENTS,
    OFD_ALLOWED_DOCUMENT_TYPES,
    OFD_RETURN_SIGNS,
    OFD_SOURCE_BOTH,
    OFD_SOURCE_MAIN,
    OFD_SOURCE_SHIFT,
    find_column,
    normalize_department,
)


class OfdParseError(ValueError):
    """Ошибка разбора файла ОФД."""


def _require_column(df: pd.DataFrame, key: str) -> str:
    column = find_column(df.columns, COLUMN_MAPPINGS["ofd"][key])
    if column is None:
        aliases = ", ".join(COLUMN_MAPPINGS["ofd"][key])
        raise OfdParseError(
            f"В отчёте ОФД не найдена колонка «{key}». "
            f"Ожидается одно из: {aliases}"
        )
    return column


def _optional_column(df: pd.DataFrame, key: str) -> str | None:
    return find_column(df.columns, COLUMN_MAPPINGS["ofd"][key])


def _to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce",
    ).fillna(0.0)


def _aggregate_ofd(work: pd.DataFrame) -> pd.DataFrame:
    grouped = (
        work.groupby(["department", "date"], as_index=False)
        .agg(
            cashless=("cashless", "sum"),
            cash=("cash", "sum"),
            credit=("credit", "sum"),
            checks_in=("checks_in", "sum"),
            returns=("returns", "sum"),
        )
        .sort_values(["department", "date"])
    )
    return grouped


def parse_ofd_checks(filepath: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Разбор основного отчёта ОФД — список чеков."""
    df = pd.read_excel(filepath, engine="openpyxl", header=0)
    total_rows = len(df)
    data = df.iloc[1:].copy()

    if not find_column(data.columns, ["Тип ФД", "Тип документа"]):
        raise OfdParseError("Файл не похож на основной отчёт ОФД со списком чеков.")

    dept_col = _require_column(data, "department")
    date_col = _require_column(data, "date")
    doc_col = _require_column(data, "document_type")
    cashless_col = _require_column(data, "cashless")
    cash_col = _require_column(data, "cash")
    sign_col = _optional_column(data, "calculation_sign")
    credit_col = _optional_column(data, "credit")

    work = data.copy()
    work["department"] = work[dept_col].map(normalize_department)
    work["date"] = pd.to_datetime(work[date_col], errors="coerce").dt.normalize()
    work["document_type"] = work[doc_col].astype(str).str.strip()
    work["cashless"] = _to_numeric(work[cashless_col])
    work["cash"] = _to_numeric(work[cash_col])
    work["credit"] = _to_numeric(work[credit_col]) if credit_col else 0.0

    work = work[work["document_type"].isin(OFD_ALLOWED_DOCUMENT_TYPES)]
    work = work[~work["department"].isin(EXCLUDED_DEPARTMENTS)]
    work = work[work["department"].notna()]
    work = work[work["date"].notna()]

    if sign_col:
        sign = work[sign_col].astype(str).str.strip()
        is_return = sign.isin(OFD_RETURN_SIGNS)
        work.loc[is_return, ["cashless", "cash", "credit"]] *= -1
        work["checks_in"] = (sign == "Приход").astype(int)
        work["returns"] = is_return.astype(int)
    else:
        work["checks_in"] = 1
        work["returns"] = 0

    grouped = _aggregate_ofd(work)
    grouped["source"] = OFD_SOURCE_MAIN

    summary = pd.DataFrame(
        {
            "Показатель": [
                "Всего строк в файле",
                "Строк после фильтрации",
                "Уникальных подразделений",
                "Период с",
                "Период по",
                "Сумма безналичных",
                "Сумма наличных",
            ],
            "Значение": [
                total_rows,
                len(work),
                grouped["department"].nunique(),
                grouped["date"].min().date() if not grouped.empty else "",
                grouped["date"].max().date() if not grouped.empty else "",
                round(grouped["cashless"].sum(), 2),
                round(grouped["cash"].sum(), 2),
            ],
        }
    )
    return summary, grouped


def parse_ofd_shifts(filepath: str) -> pd.DataFrame:
    """Разбор дополнительного отчёта ОФД — сводка по сменам."""
    df = None
    for sheet in ("Лист_1", "Sheet1"):
        try:
            candidate = pd.read_excel(filepath, sheet_name=sheet, engine="openpyxl", header=1)
        except ValueError:
            continue
        if find_column(candidate.columns, ["Название кассы"]):
            df = candidate
            break

    if df is None:
        raise OfdParseError("Файл не похож на отчёт ОФД по сменам.")

    dept_col = _require_column(df, "department")
    date_col = _require_column(df, "date")
    cashless_col = _require_column(df, "cashless")
    cash_col = _require_column(df, "cash")
    credit_col = _optional_column(df, "credit")
    checks_col = _optional_column(df, "checks_count")

    work = df.copy()
    work["department"] = work[dept_col].map(normalize_department)
    work["date"] = pd.to_datetime(work[date_col], errors="coerce").dt.normalize()
    work["cashless"] = _to_numeric(work[cashless_col])
    work["cash"] = _to_numeric(work[cash_col])
    work["credit"] = _to_numeric(work[credit_col]) if credit_col else 0.0
    work["checks_in"] = _to_numeric(work[checks_col]).astype(int) if checks_col else 0
    work["returns"] = 0

    return_col = find_column(work.columns, ["Сумма расчета наличными.1"])
    return_cashless_col = find_column(
        work.columns, ["Сумма расчета безналичными (эквайринг).1"]
    )
    if return_col:
        work["cash"] -= _to_numeric(work[return_col])
    if return_cashless_col:
        work["cashless"] -= _to_numeric(work[return_cashless_col])

    work = work[~work["department"].isin(EXCLUDED_DEPARTMENTS)]
    work = work[work["department"].notna()]
    work = work[work["date"].notna()]
    work = work[work["department"].astype(str) != "Итого:"]

    grouped = _aggregate_ofd(work)
    grouped["source"] = OFD_SOURCE_SHIFT
    return grouped


def combine_ofd_sources(main: pd.DataFrame, shift: pd.DataFrame) -> pd.DataFrame:
    """Объединить основной и дополнительный отчёты ОФД по правилам эталона."""
    main_indexed = main.set_index(["department", "date"])
    shift_indexed = shift.set_index(["department", "date"])
    keys = main_indexed.index.union(shift_indexed.index)

    rows: list[dict[str, object]] = []
    for department, date in sorted(keys):
        has_main = (department, date) in main_indexed.index
        has_shift = (department, date) in shift_indexed.index

        if has_main and has_shift:
            main_row = main_indexed.loc[(department, date)]
            shift_row = shift_indexed.loc[(department, date)]
            if isinstance(main_row, pd.DataFrame):
                main_row = main_row.iloc[0]
            if isinstance(shift_row, pd.DataFrame):
                shift_row = shift_row.iloc[0]
            rows.append(
                {
                    "department": department,
                    "date": date,
                    "cashless": float(shift_row["cashless"]),
                    "cash": float(shift_row["cash"]),
                    "credit": float(main_row["credit"]) + float(shift_row["credit"]),
                    "checks_in": int(main_row["checks_in"]) + int(shift_row["checks_in"]),
                    "returns": int(main_row["returns"]) + int(shift_row["returns"]),
                    "source": OFD_SOURCE_BOTH,
                }
            )
        elif has_shift:
            shift_row = shift_indexed.loc[(department, date)]
            if isinstance(shift_row, pd.DataFrame):
                shift_row = shift_row.iloc[0]
            rows.append(
                {
                    "department": department,
                    "date": date,
                    "cashless": float(shift_row["cashless"]),
                    "cash": float(shift_row["cash"]),
                    "credit": float(shift_row["credit"]),
                    "checks_in": int(shift_row["checks_in"]),
                    "returns": int(shift_row["returns"]),
                    "source": OFD_SOURCE_SHIFT,
                }
            )
        else:
            main_row = main_indexed.loc[(department, date)]
            if isinstance(main_row, pd.DataFrame):
                main_row = main_row.iloc[0]
            rows.append(
                {
                    "department": department,
                    "date": date,
                    "cashless": float(main_row["cashless"]),
                    "cash": float(main_row["cash"]),
                    "credit": float(main_row["credit"]),
                    "checks_in": int(main_row["checks_in"]),
                    "returns": int(main_row["returns"]),
                    "source": OFD_SOURCE_MAIN,
                }
            )

    return pd.DataFrame(rows).sort_values(["department", "date"]).reset_index(drop=True)


def parse_ofd(filepath: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Сохранён для обратной совместимости — только основной отчёт."""
    return parse_ofd_checks(filepath)

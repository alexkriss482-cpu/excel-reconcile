"""Парсер реестра кассовых документов."""

from __future__ import annotations

import pandas as pd

from config import (
    CASH_ALLOWED_DOCUMENT_TYPES,
    COLUMN_MAPPINGS,
    EXCLUDED_DEPARTMENTS,
    find_column,
    normalize_department,
)


class CashParseError(ValueError):
    """Ошибка разбора реестра кассовых документов."""


def _require_column(df: pd.DataFrame, key: str) -> str:
    column = find_column(df.columns, COLUMN_MAPPINGS["cash"][key])
    if column is None:
        aliases = ", ".join(COLUMN_MAPPINGS["cash"][key])
        raise CashParseError(
            f"В реестре кассовых документов не найдена колонка «{key}». "
            f"Ожидается одно из: {aliases}"
        )
    return column


def _to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce",
    )


def _load_registry(filepath: str) -> tuple[pd.DataFrame, int]:
    preview = pd.read_excel(filepath, engine="openpyxl", header=None)
    header_row = None
    for row_idx in range(min(20, len(preview))):
        values = {
            str(value).strip()
            for value in preview.iloc[row_idx].tolist()
            if pd.notna(value) and str(value).strip()
        }
        if "Дата" in values and "Сумма" in values:
            header_row = row_idx
            break

    if header_row is None:
        raise CashParseError("Не найдена строка заголовков с колонками «Дата» и «Сумма».")

    df = pd.read_excel(filepath, engine="openpyxl", header=header_row)
    return df, len(df)


def parse_cash(filepath: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Загрузить и обработать реестр кассовых документов.

    Возвращает:
        raw_summary — сводка по загрузке
        grouped — агрегат по подразделению и дате (amount)
    """
    df, total_rows = _load_registry(filepath)

    dept_col = _require_column(df, "department")
    date_col = _require_column(df, "date")
    amount_col = _require_column(df, "amount")
    doc_col = find_column(df.columns, COLUMN_MAPPINGS["cash"]["document_type"])

    work = df.copy()
    if doc_col:
        work = work[work[doc_col].astype(str).str.strip().isin(CASH_ALLOWED_DOCUMENT_TYPES)]

    work["department"] = work[dept_col].map(normalize_department)
    work["date"] = pd.to_datetime(work[date_col], errors="coerce", dayfirst=True).dt.normalize()
    work["amount"] = _to_numeric(work[amount_col])

    work = work[work["department"].notna()]
    work = work[work["date"].notna()]
    work = work[work["amount"].notna()]
    work = work[~work["department"].isin(EXCLUDED_DEPARTMENTS)]

    raw_summary = pd.DataFrame(
        {
            "Показатель": [
                "Всего строк в файле",
                "Строк после фильтрации",
                "Уникальных подразделений",
                "Период с",
                "Период по",
                "Сумма документов",
            ],
            "Значение": [
                total_rows,
                len(work),
                work["department"].nunique(),
                work["date"].min().date() if not work.empty else "",
                work["date"].max().date() if not work.empty else "",
                round(work["amount"].sum(), 2),
            ],
        }
    )

    grouped = (
        work.groupby(["department", "date"], as_index=False)["amount"]
        .sum()
        .sort_values(["department", "date"])
    )

    return raw_summary, grouped

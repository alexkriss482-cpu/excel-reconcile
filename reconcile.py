"""Логика сверки данных ОФД, реестра карт и кассовых документов."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from config import AMOUNT_TOLERANCE


@dataclass
class ReconcileResult:
    """Результат сверки для формирования Excel-отчёта."""

    daily: pd.DataFrame
    totals: pd.DataFrame
    discrepancies: pd.DataFrame


def _amounts_equal(left: float, right: float) -> bool:
    return abs(left - right) <= AMOUNT_TOLERANCE


def _status(match: bool) -> str:
    return "ОК" if match else "Требует проверки"


def _build_daily(ofd: pd.DataFrame, card: pd.DataFrame, cash: pd.DataFrame) -> pd.DataFrame:
    merged = ofd.merge(
        card.rename(columns={"amount": "card_amount"}),
        on=["department", "date"],
        how="outer",
    ).merge(
        cash.rename(columns={"amount": "cash_amount"}),
        on=["department", "date"],
        how="outer",
    )

    for column in ("cashless", "cash", "credit", "checks_in", "returns", "source"):
        if column not in merged.columns:
            merged[column] = 0 if column != "source" else ""

    merged["ofd_cashless"] = merged["cashless"].fillna(0.0)
    merged["ofd_cash"] = merged["cash"].fillna(0.0)
    merged["card_amount"] = merged["card_amount"].fillna(0.0)
    merged["cash_amount"] = merged["cash_amount"].fillna(0.0)
    merged["ofd_credit"] = merged["credit"].fillna(0.0)
    merged["checks_in"] = merged["checks_in"].fillna(0).astype(int)
    merged["returns"] = merged["returns"].fillna(0).astype(int)
    merged["source"] = merged["source"].fillna("")

    merged["cashless_diff"] = merged["ofd_cashless"] - merged["card_amount"]
    merged["cash_diff"] = merged["ofd_cash"] - merged["cash_amount"]
    merged["cashless_match"] = merged.apply(
        lambda row: _amounts_equal(row["ofd_cashless"], row["card_amount"]),
        axis=1,
    )
    merged["cash_match"] = merged.apply(
        lambda row: _amounts_equal(row["ofd_cash"], row["cash_amount"]),
        axis=1,
    )
    merged["has_discrepancy"] = ~(merged["cashless_match"] & merged["cash_match"])

    rows: list[dict[str, object]] = []
    for department in sorted(merged["department"].dropna().unique()):
        dept_rows = merged[merged["department"] == department].sort_values("date")
        for _, row in dept_rows.iterrows():
            rows.append(
                {
                    "Подразделение": department,
                    "Дата": row["date"],
                    "ОФД безнал": round(row["ofd_cashless"], 2),
                    "Реестр карт": round(row["card_amount"], 2),
                    "Разница безнал": round(row["cashless_diff"], 2),
                    "Статус безнал": _status(row["cashless_match"]),
                    "ОФД наличные": round(row["ofd_cash"], 2),
                    "Реестр кассовых": round(row["cash_amount"], 2),
                    "Разница наличные": round(row["cash_diff"], 2),
                    "Статус наличные": _status(row["cash_match"]),
                    "Чеков приход": int(row["checks_in"]),
                    "Возвратов": int(row["returns"]),
                    "ОФД кредит": round(row["ofd_credit"], 2),
                    "Источник ОФД": row["source"],
                    "cashless_match": row["cashless_match"],
                    "cash_match": row["cash_match"],
                    "has_discrepancy": row["has_discrepancy"],
                    "is_subtotal": False,
                }
            )

        subtotal = dept_rows[
            [
                "ofd_cashless",
                "card_amount",
                "cashless_diff",
                "ofd_cash",
                "cash_amount",
                "cash_diff",
                "checks_in",
                "returns",
                "ofd_credit",
                "cashless_match",
                "cash_match",
            ]
        ].sum(numeric_only=False)
        cashless_ok = _amounts_equal(subtotal["ofd_cashless"], subtotal["card_amount"])
        cash_ok = _amounts_equal(subtotal["ofd_cash"], subtotal["cash_amount"])
        rows.append(
            {
                "Подразделение": f"{department} ИТОГО",
                "Дата": pd.NaT,
                "ОФД безнал": round(subtotal["ofd_cashless"], 2),
                "Реестр карт": round(subtotal["card_amount"], 2),
                "Разница безнал": round(subtotal["cashless_diff"], 2),
                "Статус безнал": _status(cashless_ok),
                "ОФД наличные": round(subtotal["ofd_cash"], 2),
                "Реестр кассовых": round(subtotal["cash_amount"], 2),
                "Разница наличные": round(subtotal["cash_diff"], 2),
                "Статус наличные": _status(cash_ok),
                "Чеков приход": int(subtotal["checks_in"]),
                "Возвратов": int(subtotal["returns"]),
                "ОФД кредит": round(subtotal["ofd_credit"], 2),
                "Источник ОФД": "",
                "cashless_match": cashless_ok,
                "cash_match": cash_ok,
                "has_discrepancy": not (cashless_ok and cash_ok),
                "is_subtotal": True,
            }
        )

    return pd.DataFrame(rows)


def _build_totals(daily: pd.DataFrame) -> pd.DataFrame:
    data = daily[~daily["is_subtotal"]].copy()
    numeric_cols = [
        "ОФД безнал",
        "Реестр карт",
        "Разница безнал",
        "ОФД наличные",
        "Реестр кассовых",
        "Разница наличные",
        "ОФД кредит",
    ]

    totals = (
        data.groupby("Подразделение", as_index=False)[numeric_cols]
        .sum()
        .sort_values("Подразделение")
    )

    totals["Статус безнал"] = totals.apply(
        lambda row: _status(_amounts_equal(row["ОФД безнал"], row["Реестр карт"])),
        axis=1,
    )
    totals["Статус наличные"] = totals.apply(
        lambda row: _status(_amounts_equal(row["ОФД наличные"], row["Реестр кассовых"])),
        axis=1,
    )
    totals["Общий статус"] = totals.apply(
        lambda row: _status(row["Статус безнал"] == "ОК" and row["Статус наличные"] == "ОК"),
        axis=1,
    )
    totals["cashless_match"] = totals["Статус безнал"] == "ОК"
    totals["cash_match"] = totals["Статус наличные"] == "ОК"
    totals["has_discrepancy"] = totals["Общий статус"] != "ОК"
    return totals


def reconcile(
    ofd_grouped: pd.DataFrame,
    card_grouped: pd.DataFrame,
    cash_grouped: pd.DataFrame,
) -> ReconcileResult:
    """Выполнить сверку и подготовить данные для отчёта."""
    daily = _build_daily(ofd_grouped, card_grouped, cash_grouped)
    totals = _build_totals(daily)
    discrepancies = daily[(~daily["is_subtotal"]) & daily["has_discrepancy"]].copy()
    return ReconcileResult(daily=daily, totals=totals, discrepancies=discrepancies)

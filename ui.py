"""Графический интерфейс приложения сверки."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from excel_report import generate_report
from parser_card import parse_card
from parser_cash import parse_cash
from parser_ofd import combine_ofd_sources, parse_ofd_checks, parse_ofd_shifts
from reconcile import reconcile


class ReconcileApp(tk.Tk):
    """Главное окно приложения."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Сверка Excel: ОФД / Карты / Касса")
        self.geometry("820x380")
        self.resizable(False, False)

        self.ofd_main_path = tk.StringVar()
        self.ofd_shift_path = tk.StringVar()
        self.card_path = tk.StringVar()
        self.cash_path = tk.StringVar()
        self.status_text = tk.StringVar(
            value="Выберите файлы и нажмите «Сформировать сверку»."
        )

        self._build_ui()

    def _build_ui(self) -> None:
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=16, pady=16)

        ttk.Label(
            main,
            text="Сверка ОФД с реестрами",
            font=("", 14, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        self._add_file_row(
            main,
            row=1,
            label="1. Основной отчет ОФД (список чеков)",
            variable=self.ofd_main_path,
        )
        self._add_file_row(
            main,
            row=2,
            label="2. Доп. отчет ОФД по сменам (необязательно)",
            variable=self.ofd_shift_path,
        )
        self._add_file_row(
            main,
            row=3,
            label="3. Реестр оплат платежной картой",
            variable=self.card_path,
        )
        self._add_file_row(
            main,
            row=4,
            label="4. Реестр кассовых документов",
            variable=self.cash_path,
        )

        ttk.Button(
            main,
            text="Сформировать сверку",
            command=self._run_reconcile,
        ).grid(row=5, column=0, columnspan=3, pady=(16, 8), sticky="ew")

        ttk.Label(main, textvariable=self.status_text, wraplength=760).grid(
            row=6, column=0, columnspan=3, sticky="w"
        )

        for col in range(3):
            main.grid_columnconfigure(col, weight=1 if col == 1 else 0)

    def _add_file_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
    ) -> None:
        ttk.Label(parent, text=label, width=38).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=4
        )
        ttk.Entry(parent, textvariable=variable, width=54).grid(
            row=row, column=1, sticky="ew", pady=4
        )
        ttk.Button(
            parent,
            text="Выбрать…",
            command=lambda var=variable: self._choose_file(var),
        ).grid(row=row, column=2, padx=(8, 0), pady=4)

    def _choose_file(self, variable: tk.StringVar) -> None:
        filepath = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if filepath:
            variable.set(filepath)

    def _validate_inputs(self) -> bool:
        missing = []
        if not self.ofd_main_path.get().strip():
            missing.append("Основной отчет ОФД")
        if not self.card_path.get().strip():
            missing.append("Реестр оплат платежной картой")
        if not self.cash_path.get().strip():
            missing.append("Реестр кассовых документов")

        if missing:
            messagebox.showwarning(
                "Не хватает файлов",
                "Выберите файлы:\n• " + "\n• ".join(missing),
            )
            return False
        return True

    def _run_reconcile(self) -> None:
        if not self._validate_inputs():
            return

        ofd_main = self.ofd_main_path.get().strip()
        ofd_shift = self.ofd_shift_path.get().strip()
        card_file = self.card_path.get().strip()
        cash_file = self.cash_path.get().strip()

        output_path = filedialog.asksaveasfilename(
            title="Сохранить отчёт сверки",
            defaultextension=".xlsx",
            initialfile="сверка.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return

        self.status_text.set("Выполняется сверка…")
        self.update_idletasks()

        try:
            _, ofd_main_grouped = parse_ofd_checks(ofd_main)
            if ofd_shift:
                ofd_shift_grouped = parse_ofd_shifts(ofd_shift)
                ofd_grouped = combine_ofd_sources(ofd_main_grouped, ofd_shift_grouped)
            else:
                ofd_grouped = ofd_main_grouped

            _, card_grouped = parse_card(card_file)
            _, cash_grouped = parse_cash(cash_file)
            result = reconcile(ofd_grouped, card_grouped, cash_grouped)
            saved_path = generate_report(result, output_path)
            discrepancy_count = len(result.discrepancies)

            self.status_text.set(
                f"Готово. Отчёт сохранён: {saved_path}. "
                f"Строк с расхождениями: {discrepancy_count}."
            )
            messagebox.showinfo(
                "Сверка завершена",
                f"Отчёт сохранён:\n{saved_path}\n\n"
                f"Строк с расхождениями: {discrepancy_count}",
            )
        except Exception as exc:
            self.status_text.set("Ошибка при формировании сверки.")
            messagebox.showerror("Ошибка", str(exc))


def run_app() -> None:
    """Запустить приложение."""
    app = ReconcileApp()
    app.mainloop()

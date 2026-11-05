from __future__ import annotations

import math
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

from excel_checker import SCORING_PARTS, check_folder


class ExcelCheckerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel Checker")
        self.root.geometry("850x760")
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()

        self.reference = tk.StringVar()
        self.students = tk.StringVar()
        self.output = tk.StringVar()
        self.status = tk.StringVar(value="Kies de modeloplossing en de map met leerlingbestanden.")
        self.scoring_vars: dict[str, dict[str, tk.StringVar]] = {}

        frame = ttk.Frame(root, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Excel Checker", font=("", 20, "bold")).pack(anchor="w")
        ttk.Label(
            frame,
            text="Vergelijk .xlsx-bestanden en maak een helder HTML-rapport per leerling.",
            wraplength=620,
        ).pack(anchor="w", pady=(4, 18))

        self._file_row(frame, "Modeloplossing (.xlsx)", self.reference, self._choose_reference)
        self._file_row(frame, "Map met leerlingbestanden", self.students, self._choose_students)
        self._file_row(frame, "Map voor rapporten", self.output, self._choose_output)

        scoring_frame = ttk.LabelFrame(frame, text="Puntenverdeling (maximum per onderdeel)")
        scoring_frame.pack(fill="both", expand=True, pady=(12, 4))
        ttk.Label(
            scoring_frame,
            text="Laat een maximum leeg om dat onderdeel niet te beoordelen of in het rapport te tonen.",
            wraplength=760,
        ).pack(anchor="w", padx=8, pady=(6, 4))
        scoring_container = ttk.Frame(scoring_frame)
        scoring_container.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        canvas = tk.Canvas(scoring_container, height=235, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scoring_container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        scoring_rows = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=scoring_rows, anchor="nw")
        scoring_rows.bind(
            "<Configure>",
            lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(canvas_window, width=event.width),
        )
        self.scoring_rows = scoring_rows
        self._render_scoring_inputs([])

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(16, 8))
        self.run_button = ttk.Button(buttons, text="Start vergelijken", command=self._start)
        self.run_button.pack(side="left")
        self.progress = ttk.Progressbar(buttons, mode="indeterminate", length=180)
        self.progress.pack(side="left", padx=12)

        ttk.Separator(frame).pack(fill="x", pady=8)
        ttk.Label(frame, textvariable=self.status, wraplength=620).pack(anchor="w")
        ttk.Label(
            frame,
            text="Controleert celinhoud en formules, opmaak, werkbladen, tabellen, "
                 "validaties, voorwaardelijke opmaak en grafieken. De positie van "
                 "grafieken op het werkblad telt niet mee.",
            wraplength=620,
            foreground="#555",
        ).pack(anchor="w", pady=(12, 0))
        self.root.after(100, self._poll_messages)

    def _file_row(self, parent: ttk.Frame, label: str, variable: tk.StringVar, command: object) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text=label, width=25).pack(side="left", anchor="w")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Bladeren…", command=command).pack(side="right")

    def _choose_reference(self) -> None:
        path = filedialog.askopenfilename(
            title="Kies de modeloplossing", filetypes=[("Excel-bestand", "*.xlsx")]
        )
        if path:
            self.reference.set(path)
            try:
                sheet_names = self._reference_sheet_names(path)
            except Exception as exc:
                self._render_scoring_inputs([])
                messagebox.showerror(
                    "Modeloplossing niet leesbaar",
                    f"De werkbladen konden niet worden ingelezen:\n{exc}",
                )
                return
            self._render_scoring_inputs(sheet_names)

    @staticmethod
    def _reference_sheet_names(path: Path) -> list[str]:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            return list(workbook.sheetnames)
        finally:
            workbook.close()

    def _render_scoring_inputs(self, sheet_names: list[str]) -> None:
        previous_values = {
            (sheet, part): variable.get()
            for sheet, parts in self.scoring_vars.items()
            for part, variable in parts.items()
        }
        for child in self.scoring_rows.winfo_children():
            child.destroy()

        self.scoring_vars = {
            sheet: {
                part: tk.StringVar(
                    master=self.root,
                    value=previous_values.get((sheet, part), ""),
                )
                for part in SCORING_PARTS
            }
            for sheet in sheet_names
        }
        if not sheet_names:
            ttk.Label(
                self.scoring_rows,
                text="Kies eerst de modeloplossing om de werkbladen en puntenvelden te laden.",
                wraplength=740,
            ).grid(row=0, column=0, sticky="w", pady=4)
            return

        row = 0
        for sheet, parts in self.scoring_vars.items():
            ttk.Label(
                self.scoring_rows,
                text=sheet,
                font=("", 11, "bold"),
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(8, 3))
            row += 1
            for part, variable in parts.items():
                ttk.Label(self.scoring_rows, text=part).grid(
                    row=row, column=0, sticky="w", padx=(12, 12), pady=2,
                )
                ttk.Entry(
                    self.scoring_rows,
                    textvariable=variable,
                    width=12,
                ).grid(row=row, column=1, sticky="w", padx=(0, 14), pady=2)
                row += 1
        self.scoring_rows.columnconfigure(0, weight=1)

    def _choose_students(self) -> None:
        path = filedialog.askdirectory(title="Kies de map met leerlingbestanden")
        if path:
            self.students.set(path)
            if not self.output.get():
                self.output.set(str(Path(path) / "Excel_Checker_rapporten"))

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(title="Kies de map voor de rapporten")
        if path:
            self.output.set(path)

    def _start(self) -> None:
        reference = Path(self.reference.get())
        students = Path(self.students.get())
        output = Path(self.output.get())
        if not reference.is_file() or reference.suffix.lower() != ".xlsx":
            messagebox.showerror("Ongeldige modeloplossing", "Kies een bestaand .xlsx-bestand.")
            return
        if not students.is_dir():
            messagebox.showerror("Ongeldige map", "Kies een bestaande map met leerlingbestanden.")
            return
        if not self.output.get().strip():
            messagebox.showerror("Geen rapportmap", "Kies een map om de rapporten in op te slaan.")
            return
        try:
            actual_sheet_names = self._reference_sheet_names(reference)
            if actual_sheet_names != list(self.scoring_vars):
                self._render_scoring_inputs(actual_sheet_names)
                messagebox.showerror(
                    "Modeloplossing gewijzigd",
                    "De werkbladen zijn opnieuw ingelezen. Controleer de puntenverdeling en start opnieuw.",
                )
                return
            scoring_parts = self._read_scoring_parts()
        except ValueError as exc:
            messagebox.showerror("Ongeldige puntenverdeling", str(exc))
            return
        except Exception as exc:
            messagebox.showerror(
                "Modeloplossing niet leesbaar",
                f"De werkbladen konden niet worden ingelezen:\n{exc}",
            )
            return
        self.run_button.configure(state="disabled")
        self.progress.start(12)
        self.status.set("Bestanden worden vergeleken…")
        threading.Thread(
            target=self._run_comparison,
            args=(reference, students, output, scoring_parts),
            daemon=True,
        ).start()

    def _read_scoring_parts(self) -> dict[str, dict[str, float]]:
        configured: dict[str, dict[str, float]] = {}
        for sheet, parts in self.scoring_vars.items():
            sheet_parts = {}
            for part, variable in parts.items():
                raw_value = variable.get().strip()
                if not raw_value:
                    continue
                try:
                    maximum = float(raw_value.replace(",", "."))
                except ValueError as exc:
                    raise ValueError(
                        f"Vul een geldig maximum in voor '{part}' op werkblad '{sheet}'."
                    ) from exc
                if not math.isfinite(maximum) or maximum <= 0:
                    raise ValueError(
                        f"Het maximum voor '{part}' op werkblad '{sheet}' moet groter zijn dan 0."
                    )
                sheet_parts[part] = maximum
            configured[sheet] = sheet_parts
        return configured

    def _run_comparison(
        self,
        reference: Path,
        students: Path,
        output: Path,
        scoring_parts: dict[str, dict[str, float]],
    ) -> None:
        try:
            reports = check_folder(reference, students, output, scoring_parts)
            self.messages.put(("done", (reports, output)))
        except Exception as exc:
            self.messages.put(("error", str(exc)))

    def _poll_messages(self) -> None:
        try:
            kind, payload = self.messages.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_messages)
            return
        self.progress.stop()
        self.run_button.configure(state="normal")
        if kind == "error":
            self.status.set("Het vergelijken is niet gelukt.")
            messagebox.showerror("Fout bij vergelijken", str(payload))
        else:
            reports, output = payload
            if reports:
                self.status.set(f"{len(reports)} rapport(en) opgeslagen in:\n{output}")
            else:
                self.status.set(f"Geen .xlsx-leerlingbestanden gevonden in:\n{self.students.get()}")
        self.root.after(100, self._poll_messages)


def main() -> None:
    root = tk.Tk()
    ExcelCheckerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

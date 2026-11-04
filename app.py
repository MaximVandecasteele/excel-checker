from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from excel_checker import check_folder


class ExcelCheckerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel Checker")
        self.root.geometry("680x430")
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()

        self.reference = tk.StringVar()
        self.students = tk.StringVar()
        self.output = tk.StringVar()
        self.status = tk.StringVar(value="Kies de modeloplossing en de map met leerlingbestanden.")

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
        self.run_button.configure(state="disabled")
        self.progress.start(12)
        self.status.set("Bestanden worden vergeleken…")
        threading.Thread(
            target=self._run_comparison, args=(reference, students, output), daemon=True
        ).start()

    def _run_comparison(self, reference: Path, students: Path, output: Path) -> None:
        try:
            reports = check_folder(reference, students, output)
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

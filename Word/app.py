from __future__ import annotations

import math
import queue
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from Shared.pdf_exporter import convert_html_reports, find_html_reports
from Shared.report_server import make_report_server
from Word.word_checker import (
    check_folder,
    reference_style_names,
    scoring_part_names,
)


class WordCheckerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Word Checker")
        self.root.geometry("860x760")
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()
        self.report_server = None
        self.report_server_thread: threading.Thread | None = None
        self.reference = tk.StringVar()
        self.students = tk.StringVar()
        self.filename_suffix = tk.StringVar()
        self.output = tk.StringVar()
        self.status = tk.StringVar(value="Kies de modeloplossing en de map met leerlingbestanden.")
        self.scoring_vars: dict[str, tk.StringVar] = {}

        frame = ttk.Frame(root, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Word Checker", font=("", 20, "bold")).pack(anchor="w")
        ttk.Label(
            frame,
            text="Vergelijk .docx-bestanden met de modeloplossing en maak een HTML-rapport per leerling.",
            wraplength=700,
        ).pack(anchor="w", pady=(4, 14))
        self._file_row(frame, "Modeloplossing (.docx)", self.reference, self._choose_reference)
        self._file_row(frame, "Map met leerlingbestanden", self.students, self._choose_students)

        suffix_row = ttk.Frame(frame)
        suffix_row.pack(fill="x", pady=6)
        ttk.Label(suffix_row, text="Bestandsnaam eindigt op", width=25).pack(side="left", anchor="w")
        ttk.Entry(suffix_row, textvariable=self.filename_suffix).pack(
            side="left", fill="x", expand=True, padx=8,
        )
        ttk.Label(suffix_row, text="(.docx wordt toegevoegd)").pack(side="right")
        ttk.Label(
            frame,
            text="Spaties en hoofdletters worden genegeerd. Laat leeg om alle .docx-bestanden te verwerken.",
            wraplength=760,
            foreground="#555",
        ).pack(anchor="w", pady=(0, 4))
        self._file_row(frame, "Map voor rapporten", self.output, self._choose_output)

        scoring_frame = ttk.LabelFrame(frame, text="Puntenverdeling (maximum per stijl/controle)")
        scoring_frame.pack(fill="both", expand=True, pady=(10, 4))
        ttk.Label(
            scoring_frame,
            text="Vul punten in voor onderdelen die je wilt beoordelen. Een onderdeel zonder verschillen krijgt automatisch de maximumscore.",
            wraplength=760,
        ).pack(anchor="w", padx=8, pady=(6, 4))
        container = ttk.Frame(scoring_frame)
        container.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        canvas = tk.Canvas(container, height=235, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.scoring_rows = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=self.scoring_rows, anchor="nw")
        self.scoring_rows.bind(
            "<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.bind(
            "<Configure>", lambda event: canvas.itemconfigure(canvas_window, width=event.width)
        )
        self._render_scoring_inputs([])

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(12, 8))
        self.run_button = ttk.Button(buttons, text="Start vergelijken", command=self._start)
        self.run_button.pack(side="left")
        self.progress = ttk.Progressbar(buttons, mode="indeterminate", length=180)
        self.progress.pack(side="left", padx=12)

        reports = ttk.LabelFrame(frame, text="Rapporten")
        reports.pack(fill="x", pady=(8, 4))
        self.server_button = ttk.Button(
            reports, text="Open rapporten en vul scores in", command=self._start_report_server
        )
        self.server_button.pack(side="left", padx=8, pady=8)
        self.pdf_button = ttk.Button(
            reports, text="Exporteer rapporten naar PDF", command=self._start_pdf_export
        )
        self.pdf_button.pack(side="left", padx=(0, 8), pady=8)
        ttk.Separator(frame).pack(fill="x", pady=8)
        ttk.Label(frame, textvariable=self.status, wraplength=700).pack(anchor="w")
        ttk.Label(
            frame,
            text="Controleert tekens en witruimte, gebruikte stijlen en stijldefinities, automatische nummering, inhoudsopgave en pagina-eindes.",
            wraplength=700,
            foreground="#555",
        ).pack(anchor="w", pady=(10, 0))
        self.root.after(100, self._poll_messages)
        self.root.protocol("WM_DELETE_WINDOW", self._close)

    @staticmethod
    def _file_row(parent: ttk.Frame, label: str, variable: tk.StringVar, command: object) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text=label, width=25).pack(side="left", anchor="w")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Bladeren...", command=command).pack(side="right")

    def _choose_reference(self) -> None:
        path = filedialog.askopenfilename(
            title="Kies de modeloplossing", filetypes=[("Word-bestand", "*.docx")]
        )
        if not path:
            return
        self.reference.set(path)
        try:
            self._render_scoring_inputs(reference_style_names(Path(path)))
        except Exception as exc:
            self._render_scoring_inputs([])
            messagebox.showerror("Modeloplossing niet leesbaar", f"Het bestand kon niet worden ingelezen:\n{exc}")

    def _render_scoring_inputs(self, styles: list[str]) -> None:
        previous = {name: variable.get() for name, variable in self.scoring_vars.items()}
        for child in self.scoring_rows.winfo_children():
            child.destroy()
        names = scoring_part_names(styles)
        self.scoring_vars = {
            name: tk.StringVar(master=self.root, value=previous.get(name, ""))
            for name in names
        }
        if not names:
            ttk.Label(
                self.scoring_rows,
                text="Kies eerst de modeloplossing om gebruikte stijlen en puntenvelden te laden.",
                wraplength=740,
            ).grid(row=0, column=0, sticky="w", pady=4)
            return
        for row, (name, variable) in enumerate(self.scoring_vars.items()):
            ttk.Label(self.scoring_rows, text=name).grid(
                row=row, column=0, sticky="w", padx=(0, 12), pady=2,
            )
            ttk.Entry(self.scoring_rows, textvariable=variable, width=12).grid(
                row=row, column=1, sticky="w", pady=2,
            )
        self.scoring_rows.columnconfigure(0, weight=1)

    def _choose_students(self) -> None:
        path = filedialog.askdirectory(title="Kies de map met leerlingbestanden")
        if path:
            self.students.set(path)
            if not self.output.get():
                self.output.set(str(Path(path) / "Word_Checker_rapporten"))

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(title="Kies de map voor de rapporten")
        if path:
            self.output.set(path)

    def _read_scoring_parts(self) -> dict[str, float]:
        configured: dict[str, float] = {}
        for name, variable in self.scoring_vars.items():
            raw_value = variable.get().strip()
            if not raw_value:
                continue
            try:
                maximum = float(raw_value.replace(",", "."))
            except ValueError as exc:
                raise ValueError(f"Vul een geldig maximum in voor '{name}'.") from exc
            if not math.isfinite(maximum) or maximum <= 0:
                raise ValueError(f"Het maximum voor '{name}' moet groter zijn dan 0.")
            configured[name] = maximum
        return configured

    def _start(self) -> None:
        reference = Path(self.reference.get())
        students = Path(self.students.get())
        output = Path(self.output.get())
        if not reference.is_file() or reference.suffix.lower() != ".docx":
            messagebox.showerror("Ongeldige modeloplossing", "Kies een bestaand .docx-bestand.")
            return
        if not students.is_dir():
            messagebox.showerror("Ongeldige map", "Kies een bestaande map met leerlingbestanden.")
            return
        if not self.output.get().strip():
            messagebox.showerror("Geen rapportmap", "Kies een map om de rapporten in op te slaan.")
            return
        try:
            actual_styles = reference_style_names(reference)
            expected_names = [name.removeprefix("Stijl: ") for name in self.scoring_vars if name.startswith("Stijl: ")]
            if actual_styles != expected_names:
                self._render_scoring_inputs(actual_styles)
                messagebox.showerror(
                    "Modeloplossing gewijzigd",
                    "De gebruikte stijlen zijn opnieuw ingelezen. Controleer de puntenverdeling en start opnieuw.",
                )
                return
            scoring_parts = self._read_scoring_parts()
        except ValueError as exc:
            messagebox.showerror("Ongeldige puntenverdeling", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Modeloplossing niet leesbaar", str(exc))
            return
        self.run_button.configure(state="disabled")
        self.progress.start(12)
        self.status.set("Bestanden worden vergeleken...")
        threading.Thread(
            target=self._run_comparison,
            args=(reference, students, output, scoring_parts, self.filename_suffix.get().strip()),
            daemon=True,
        ).start()

    def _run_comparison(
        self,
        reference: Path,
        students: Path,
        output: Path,
        scoring: dict[str, float],
        suffix: str,
    ) -> None:
        try:
            reports = check_folder(reference, students, output, scoring, suffix)
            self.messages.put(("done", (reports, output)))
        except Exception as exc:
            self.messages.put(("error", str(exc)))

    def _start_report_server(self) -> None:
        output = Path(self.output.get())
        if not output.is_dir() or not find_html_reports(output):
            messagebox.showerror("Geen rapporten", "Maak eerst rapporten aan met 'Start vergelijken'.")
            return
        if self.report_server is not None:
            address = f"http://127.0.0.1:{self.report_server.server_port}/"
            webbrowser.open(address)
            return
        try:
            self.report_server = make_report_server(
                output, on_shutdown=lambda: self.messages.put(("server_stopped", None))
            )
        except OSError as exc:
            messagebox.showerror("Server niet beschikbaar", str(exc))
            return
        self.report_server_thread = threading.Thread(
            target=self.report_server.serve_forever, daemon=True
        )
        self.report_server_thread.start()
        address = f"http://127.0.0.1:{self.report_server.server_port}/"
        self.server_button.configure(state="disabled")
        self.status.set(f"Rapportserver actief:\n{address}")
        if not webbrowser.open(address):
            messagebox.showinfo("Rapportserver actief", f"Open dit adres in je browser:\n{address}")

    def _start_pdf_export(self) -> None:
        source = Path(self.output.get())
        if not source.is_dir() or not find_html_reports(source):
            messagebox.showerror("Geen rapporten", "Maak eerst rapporten aan met 'Start vergelijken'.")
            return
        destination = filedialog.askdirectory(
            title="Kies de map voor de PDF-rapporten", initialdir=str(source)
        )
        if not destination:
            return
        self.pdf_button.configure(state="disabled")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        threading.Thread(
            target=self._run_pdf_export, args=(source, Path(destination)), daemon=True
        ).start()

    def _run_pdf_export(self, source: Path, destination: Path) -> None:
        try:
            paths = convert_html_reports(source, destination)
            self.messages.put(("pdf_done", (paths, destination)))
        except Exception as exc:
            self.messages.put(("pdf_error", str(exc)))

    def _poll_messages(self) -> None:
        try:
            kind, payload = self.messages.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_messages)
            return
        if kind in {"done", "error"}:
            self.progress.stop()
            self.run_button.configure(state="normal")
        if kind == "done":
            reports, output = payload
            self.status.set(
                f"{len(reports)} rapport(en) opgeslagen in:\n{output}"
                if reports else "Geen overeenkomende .docx-leerlingbestanden gevonden."
            )
        elif kind == "error":
            self.status.set("Het vergelijken is niet gelukt.")
            messagebox.showerror("Fout bij vergelijken", str(payload))
        elif kind == "pdf_done":
            self.progress.stop()
            self.pdf_button.configure(state="normal")
            self.status.set(f"{len(payload[0])} PDF-bestand(en) opgeslagen in:\n{payload[1]}")
        elif kind == "pdf_error":
            self.progress.stop()
            self.pdf_button.configure(state="normal")
            messagebox.showerror("Fout bij PDF-export", str(payload))
        elif kind == "server_stopped":
            self.report_server = None
            self.report_server_thread = None
            self.server_button.configure(state="normal")
        self.root.after(100, self._poll_messages)

    def _close(self) -> None:
        if self.report_server is not None:
            self.report_server.shutdown()
            self.report_server.server_close()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    WordCheckerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

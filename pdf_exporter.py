from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable


def find_html_reports(folder: Path) -> list[Path]:
    return sorted(
        path for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() in {".html", ".htm"}
    )


def convert_html_reports(
    html_folder: Path,
    output_folder: Path,
    on_progress: Callable[[int, int, Path], None] | None = None,
) -> list[Path]:
    reports = find_html_reports(html_folder)
    if not reports:
        return []

    output_folder.mkdir(parents=True, exist_ok=True)
    from playwright.sync_api import sync_playwright

    pdf_paths = []
    used_stems: dict[str, int] = {}
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        try:
            page = browser.new_page()
            for index, report in enumerate(reports, start=1):
                page.goto(report.resolve().as_uri(), wait_until="load")
                page.evaluate("document.fonts.ready")
                used_stems[report.stem] = used_stems.get(report.stem, 0) + 1
                suffix = f"_{used_stems[report.stem]}" if used_stems[report.stem] > 1 else ""
                pdf_path = output_folder / f"{report.stem}{suffix}.pdf"
                page.pdf(
                    path=str(pdf_path),
                    format="A4",
                    print_background=True,
                    prefer_css_page_size=True,
                    margin={
                        "top": "12mm",
                        "right": "12mm",
                        "bottom": "12mm",
                        "left": "12mm",
                    },
                )
                pdf_paths.append(pdf_path)
                if on_progress is not None:
                    on_progress(index, len(reports), report)
        finally:
            browser.close()
    return pdf_paths


class PdfExporterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel Checker - PDF-export")
        self.root.geometry("680x280")
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()

        self.source = tk.StringVar()
        self.destination = tk.StringVar()
        self.status = tk.StringVar(
            value="Kies de map met HTML-rapporten en een map voor de PDF-bestanden."
        )

        frame = ttk.Frame(root, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Rapporten omzetten naar PDF", font=("", 18, "bold")).pack(
            anchor="w", pady=(0, 12)
        )
        self._folder_row(frame, "Map met HTML-rapporten", self.source, self._choose_source)
        self._folder_row(frame, "Map voor PDF-bestanden", self.destination, self._choose_destination)

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(14, 8))
        self.convert_button = ttk.Button(
            buttons,
            text="Zet om naar PDF",
            command=self._start,
        )
        self.convert_button.pack(side="left")
        self.progress = ttk.Progressbar(buttons, mode="determinate", length=240)
        self.progress.pack(side="left", padx=12)

        ttk.Label(frame, textvariable=self.status, wraplength=620).pack(anchor="w", pady=(8, 0))
        ttk.Label(
            frame,
            text="Elke HTML krijgt een afzonderlijke PDF. De HTML-kopieën die je met "
                 "de rapportviewer hebt beoordeeld, bevatten de opgeslagen scores.",
            wraplength=620,
            foreground="#555",
        ).pack(anchor="w", pady=(10, 0))
        self.root.after(100, self._poll_messages)

    @staticmethod
    def _folder_row(
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        command: Callable[[], None],
    ) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text=label, width=25).pack(side="left", anchor="w")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Bladeren…", command=command).pack(side="right")

    def _choose_source(self) -> None:
        path = filedialog.askdirectory(title="Kies de map met opgeslagen HTML-rapporten")
        if path:
            self.source.set(path)
            if not self.destination.get():
                self.destination.set(str(Path(path) / "PDF"))

    def _choose_destination(self) -> None:
        path = filedialog.askdirectory(title="Kies de map voor de PDF-bestanden")
        if path:
            self.destination.set(path)

    def _start(self) -> None:
        source = Path(self.source.get())
        destination = Path(self.destination.get())
        if not source.is_dir():
            messagebox.showerror("Ongeldige map", "Kies een bestaande map met HTML-rapporten.")
            return
        if not self.destination.get().strip():
            messagebox.showerror("Geen uitvoermap", "Kies een map voor de PDF-bestanden.")
            return
        reports = find_html_reports(source)
        if not reports:
            messagebox.showerror("Geen HTML-rapporten", "De gekozen map bevat geen HTML-rapporten.")
            return

        self.convert_button.configure(state="disabled")
        self.progress.configure(maximum=len(reports), value=0)
        self.status.set(f"{len(reports)} rapport(en) worden naar PDF omgezet…")
        threading.Thread(
            target=self._convert,
            args=(source, destination),
            daemon=True,
        ).start()

    def _convert(self, source: Path, destination: Path) -> None:
        try:
            def report_progress(index: int, total: int, report: Path) -> None:
                self.messages.put(("progress", (index, total, report.name)))

            paths = convert_html_reports(source, destination, report_progress)
            self.messages.put(("done", (paths, destination)))
        except Exception as exc:
            self.messages.put(("error", str(exc)))

    def _poll_messages(self) -> None:
        try:
            kind, payload = self.messages.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_messages)
            return

        if kind == "progress":
            index, total, filename = payload
            self.progress.configure(value=index)
            self.status.set(f"PDF {index}/{total}: {filename}")
        elif kind == "error":
            self.convert_button.configure(state="normal")
            self.status.set("Het omzetten naar PDF is niet gelukt.")
            messagebox.showerror(
                "Fout bij PDF-export",
                f"{payload}\n\nControleer of Chromium is geïnstalleerd met:\n"
                "python3 -m playwright install chromium",
            )
        else:
            self.convert_button.configure(state="normal")
            paths, destination = payload
            self.progress.configure(value=len(paths))
            self.status.set(f"{len(paths)} PDF-bestand(en) opgeslagen in:\n{destination}")
        self.root.after(100, self._poll_messages)


def main() -> None:
    root = tk.Tk()
    PdfExporterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

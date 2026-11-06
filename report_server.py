from __future__ import annotations

import json
import math
import queue
import re
import threading
import tkinter as tk
import webbrowser
from html import escape
from html.parser import HTMLParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib.parse import quote, unquote, urlsplit
from typing import Any, Callable


_MAX_REQUEST_BYTES = 64 * 1024
_SCORE_DIRECTORY = ".excel_checker_scores"


class _ScoreInputParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.maxima: list[float] = []
        self.defaults: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag != "input":
            return
        attributes = dict(attrs)
        if attributes.get("data-score-field") != "earned":
            return
        try:
            maximum = float(attributes["data-maximum"] or "")
        except (KeyError, TypeError, ValueError):
            return
        if math.isfinite(maximum) and maximum > 0:
            self.maxima.append(maximum)
            self.defaults.append(attributes.get("value") or "")


def _read_report_maxima(report_path: Path) -> list[float]:
    parser = _ScoreInputParser()
    parser.feed(report_path.read_text(encoding="utf-8"))
    return parser.maxima


def _read_report_score_defaults(report_path: Path) -> list[str]:
    parser = _ScoreInputParser()
    parser.feed(report_path.read_text(encoding="utf-8"))
    return parser.defaults


def _persist_report_scores(report_path: Path, scores: list[str]) -> None:
    report = report_path.read_text(encoding="utf-8")
    input_pattern = re.compile(
        r'<input\b(?=[^>]*\bid="score-(\d+)")[^>]*>',
        re.IGNORECASE,
    )
    updated = set()

    def update_input(match: re.Match[str]) -> str:
        index = int(match.group(1))
        if index >= len(scores):
            return match.group(0)
        tag = match.group(0)
        tag = re.sub(r'\svalue="[^"]*"', "", tag)
        if scores[index] != "":
            tag = tag[:-1] + f' value="{escape(scores[index], quote=True)}">'
        updated.add(index)
        return tag

    report = input_pattern.sub(update_input, report)
    if updated != set(range(len(scores))):
        raise ValueError("Niet alle scorevelden in het HTML-rapport konden worden bijgewerkt.")
    if re.search(r'<main\b[^>]*\bdata-scores-embedded="true"', report) is None:
        report = re.sub(
            r'(<main\b[^>]*\bclass="grading"[^>]*)(>)',
            r'\1 data-scores-embedded="true"\2',
            report,
            count=1,
        )

    temporary = report_path.with_suffix(report_path.suffix + ".tmp")
    temporary.write_text(report, encoding="utf-8")
    temporary.replace(report_path)


def _score_file(folder: Path, report_name: str) -> Path:
    return folder / _SCORE_DIRECTORY / f"{report_name}.json"


def make_report_server(
    report_folder: Path,
    host: str = "127.0.0.1",
    port: int = 0,
    on_shutdown: Callable[[], None] | None = None,
) -> ThreadingHTTPServer:
    report_folder = report_folder.resolve()
    score_lock = threading.Lock()

    def find_report(report_name: str) -> Path | None:
        if not report_name or Path(report_name).name != report_name:
            return None
        candidate = report_folder / report_name
        if (
            candidate.is_file()
            and candidate.resolve().parent == report_folder
            and candidate.suffix.lower() in {".html", ".htm"}
        ):
            return candidate
        return None

    def list_reports() -> list[Path]:
        return sorted(
            path for path in report_folder.iterdir()
            if path.is_file() and path.suffix.lower() in {".html", ".htm"}
        )

    class ReportRequestHandler(BaseHTTPRequestHandler):
        server_version = "ExcelCheckerReportServer"

        def log_message(self, _format: str, *_args: Any) -> None:
            return

        def _send(self, status: int, content_type: str, body: bytes) -> None:
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(body)

        def _send_json(self, status: int, payload: dict[str, Any]) -> None:
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self._send(status, "application/json; charset=utf-8", body)

        @staticmethod
        def _navigation(
            reports: list[Path],
            index: int,
            include_counter: bool,
        ) -> str:
            previous = reports[(index - 1) % len(reports)]
            following = reports[(index + 1) % len(reports)]
            counter = (
                f'<span class="report-counter">{index + 1} / {len(reports)}</span>'
                if include_counter else ""
            )
            return (
                '<nav class="report-navigation" aria-label="Rapportnavigatie">'
                f'<a class="report-button" href="/reports/{quote(previous.name)}">'
                "← Vorige toets</a>"
                '<a class="report-button" href="/">← Overzicht</a>'
                f'<a class="report-button" href="/reports/{quote(following.name)}">'
                f"Volgende toets →</a>"
                '<button type="button" class="report-button end-evaluation">'
                "Beëindig evaluatie</button>"
                f"{counter}</nav>"
            )

        def do_GET(self) -> None:
            request = urlsplit(self.path)
            if request.path == "/":
                reports = list_reports()
                links = "".join(
                    f'<li><a class="report-link" '
                    f'data-report-name="{escape(path.name, quote=True)}" '
                    f'href="/reports/{quote(path.name)}">'
                    f"{escape(path.stem)}</a></li>"
                    for path in reports
                )
                page = (
                    "<!doctype html><html lang=\"nl\"><meta charset=\"utf-8\">"
                    "<meta name=\"viewport\" content=\"width=device-width\">"
                    "<title>Excel Checker — rapporten</title>"
                    "<style>body{font:16px/1.5 system-ui,sans-serif;max-width:900px;"
                    "margin:2rem auto;padding:0 1rem}li{margin:.5rem 0}"
                    ".report-link.completed{color:#137333;font-weight:700}"
                    ".end-evaluation{margin:1rem 0;padding:.5rem .8rem;font:inherit;"
                    "cursor:pointer}"
                    "@media print{.end-evaluation{display:none!important}}</style>"
                    "<h1>Excel Checker — rapporten</h1>"
                    "<p>Scores worden automatisch op deze computer opgeslagen.</p>"
                    '<button type="button" class="end-evaluation">'
                    "Beëindig evaluatie</button>"
                    f"<ul>{links or '<li>Geen HTML-rapporten gevonden.</li>'}</ul>"
                    "<script>"
                    "document.querySelectorAll('.report-link').forEach(async link=>{"
                    "try{const response=await fetch('/api/scores/'+"
                    "encodeURIComponent(link.dataset.reportName));"
                    "if(!response.ok)throw new Error('Scorestatus niet beschikbaar.');"
                    "const status=await response.json();"
                    "if(status.complete)link.classList.add('completed');"
                    "}catch(error){console.error(error);}});"
                    "document.querySelectorAll('.end-evaluation').forEach(button=>{"
                    "button.addEventListener('click',async()=>{"
                    "if(!window.confirm('Wil je de evaluatie beëindigen?'))return;"
                    "button.disabled=true;"
                    "try{const response=await fetch('/api/shutdown',{method:'POST'});"
                    "if(!response.ok)throw new Error('Evaluatie beëindigen is niet gelukt.');"
                    "document.body.innerHTML='<h1>Evaluatie beëindigd</h1>';"
                    "}catch(error){button.disabled=false;window.alert(error.message);}"
                    "});});"
                    "</script>"
                    "</html>"
                ).encode("utf-8")
                self._send(200, "text/html; charset=utf-8", page)
                return

            if request.path.startswith("/reports/"):
                name = unquote(request.path[len("/reports/"):])
                report_path = find_report(name)
                if report_path is None:
                    self._send(404, "text/plain; charset=utf-8", b"Rapport niet gevonden.")
                    return
                try:
                    body = report_path.read_text(encoding="utf-8")
                except OSError:
                    self._send(500, "text/plain; charset=utf-8", b"Rapport kon niet worden gelezen.")
                    return
                reports = list_reports()
                index = reports.index(report_path)
                navigation = self._navigation(reports, index, True)
                bottom_navigation = self._navigation(reports, index, False)
                report_style = (
                    "<style>"
                    ".report-navigation{display:flex;align-items:center;gap:.6rem;"
                    "flex-wrap:wrap;margin:1rem 0 1.5rem}"
                    ".report-button{display:inline-block;padding:.45rem .75rem;"
                    "border:1px solid #dadce0;border-radius:.35rem;color:#202124;"
                    "text-decoration:none;background:#f8f9fa}"
                    ".report-button:hover{background:#e8eaed}"
                    ".report-counter{margin-left:auto;font-weight:700;white-space:nowrap}"
                    ".end-evaluation{font:inherit;cursor:pointer}"
                    "@media print{.report-navigation{display:none!important}}"
                    "</style>"
                )
                insertion = body.lower().find("<h1")
                if insertion >= 0:
                    body = body[:insertion] + report_style + navigation + body[insertion:]
                script_position = body.lower().find("<script")
                if script_position >= 0:
                    body = body[:script_position] + bottom_navigation + body[script_position:]
                else:
                    body += bottom_navigation
                body = body.replace(
                    "</html>",
                    "<script>"
                    "document.querySelectorAll('.end-evaluation').forEach(button=>{"
                    "button.addEventListener('click',async()=>{"
                    "if(!window.confirm('Wil je de evaluatie beëindigen?'))return;"
                    "button.disabled=true;"
                    "try{const response=await fetch('/api/shutdown',{method:'POST'});"
                    "if(!response.ok)throw new Error('Evaluatie beëindigen is niet gelukt.');"
                    "document.body.innerHTML='<h1>Evaluatie beëindigd</h1>';"
                    "}catch(error){button.disabled=false;window.alert(error.message);}"
                    "});});"
                    "</script></html>",
                )
                self._send(200, "text/html; charset=utf-8", body.encode("utf-8"))
                return

            if request.path.startswith("/api/scores/"):
                name = unquote(request.path[len("/api/scores/"):])
                report_path = find_report(name)
                if report_path is None:
                    self._send_json(404, {"error": "Rapport niet gevonden."})
                    return
                score_path = _score_file(report_folder, report_path.name)
                with score_lock:
                    try:
                        scores = json.loads(score_path.read_text(encoding="utf-8"))["scores"]
                    except FileNotFoundError:
                        scores = []
                    except (OSError, json.JSONDecodeError, KeyError, TypeError):
                        self._send_json(500, {"error": "Opgeslagen scores konden niet worden gelezen."})
                        return
                    try:
                        defaults = _read_report_score_defaults(report_path)
                    except (OSError, UnicodeDecodeError):
                        self._send_json(500, {"error": "Rapportscores konden niet worden gelezen."})
                        return
                completion_scores = scores if len(scores) == len(defaults) else defaults
                complete = bool(completion_scores) and all(
                    isinstance(score, str) and score.strip() != ""
                    for score in completion_scores
                )
                self._send_json(200, {"scores": scores, "complete": complete})
                return

            self._send(404, "text/plain; charset=utf-8", b"Niet gevonden.")

        def do_POST(self) -> None:
            request = urlsplit(self.path)
            if request.path == "/api/shutdown":
                origin = urlsplit(self.headers.get("Origin", ""))
                if (
                    origin.scheme != "http"
                    or origin.netloc != self.headers.get("Host")
                ):
                    self._send_json(403, {"error": "Ongeldige herkomst."})
                    return
                self._send_json(200, {"stopping": True})

                def stop_server() -> None:
                    self.server.shutdown()
                    self.server.server_close()
                    if on_shutdown is not None:
                        on_shutdown()

                threading.Thread(target=stop_server, daemon=True).start()
                return

            if not request.path.startswith("/api/scores/"):
                self._send_json(404, {"error": "Niet gevonden."})
                return

            name = unquote(request.path[len("/api/scores/"):])
            report_path = find_report(name)
            if report_path is None:
                self._send_json(404, {"error": "Rapport niet gevonden."})
                return

            try:
                content_length = int(self.headers.get("Content-Length", ""))
            except ValueError:
                self._send_json(400, {"error": "Ongeldige requestlengte."})
                return
            if content_length <= 0 or content_length > _MAX_REQUEST_BYTES:
                self._send_json(413, {"error": "Scoregegevens zijn te groot of ontbreken."})
                return

            try:
                payload = json.loads(self.rfile.read(content_length))
            except (json.JSONDecodeError, UnicodeDecodeError):
                self._send_json(400, {"error": "Ongeldige JSON-scoregegevens."})
                return
            scores = payload.get("scores") if isinstance(payload, dict) else None
            maxima = _read_report_maxima(report_path)
            if (
                not isinstance(scores, list)
                or len(scores) != len(maxima)
                or any(not isinstance(score, str) or len(score) > 64 for score in scores)
            ):
                self._send_json(400, {"error": "Scoregegevens komen niet overeen met het rapport."})
                return

            for score, maximum in zip(scores, maxima):
                if score == "":
                    continue
                try:
                    value = float(score)
                except ValueError:
                    self._send_json(400, {"error": "Een score is geen geldig getal."})
                    return
                if not math.isfinite(value) or value < 0 or value > maximum:
                    self._send_json(400, {"error": "Een score valt buiten het toegestane bereik."})
                    return

            score_path = _score_file(report_folder, report_path.name)
            with score_lock:
                try:
                    _persist_report_scores(report_path, scores)
                    score_path.parent.mkdir(parents=True, exist_ok=True)
                    temporary = score_path.with_suffix(".tmp")
                    temporary.write_text(
                        json.dumps({"scores": scores}, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    temporary.replace(score_path)
                except OSError as exc:
                    self._send_json(500, {"error": f"Scores opslaan is niet gelukt: {exc}"})
                    return
            self._send_json(200, {"saved": True})

    return ThreadingHTTPServer((host, port), ReportRequestHandler)


class ReportServerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel Checker - rapporten en scores")
        self.root.geometry("680x300")
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()
        self.report_folder = tk.StringVar()
        self.status = tk.StringVar(value="Kies de map met HTML-rapporten.")
        self.server: ThreadingHTTPServer | None = None
        self.server_thread: threading.Thread | None = None

        frame = ttk.Frame(root, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Lokale rapportviewer", font=("", 18, "bold")).pack(
            anchor="w", pady=(0, 12)
        )
        row = ttk.Frame(frame)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text="Map met HTML-rapporten", width=25).pack(side="left", anchor="w")
        ttk.Entry(row, textvariable=self.report_folder).pack(
            side="left", fill="x", expand=True, padx=8
        )
        ttk.Button(row, text="Bladeren…", command=self._choose_folder).pack(side="right")

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(16, 8))
        self.start_button = ttk.Button(buttons, text="Open rapporten", command=self._start)
        self.start_button.pack(side="left")
        ttk.Label(frame, textvariable=self.status, wraplength=620).pack(
            anchor="w", pady=(8, 0)
        )
        ttk.Label(
            frame,
            text="Laat dit venster open terwijl je scores invult. De scores worden "
                 "automatisch naast de HTML-rapporten bewaard.",
            wraplength=620,
            foreground="#555",
        ).pack(anchor="w", pady=(12, 0))
        self.root.protocol("WM_DELETE_WINDOW", self._close)

    def _choose_folder(self) -> None:
        path = filedialog.askdirectory(title="Kies de map met HTML-rapporten")
        if path:
            self.report_folder.set(path)

    def _start(self) -> None:
        folder = Path(self.report_folder.get())
        if not folder.is_dir():
            messagebox.showerror("Ongeldige map", "Kies een bestaande map met HTML-rapporten.")
            return
        reports = [
            path for path in folder.iterdir()
            if path.is_file() and path.suffix.lower() in {".html", ".htm"}
        ]
        if not reports:
            messagebox.showerror("Geen rapporten", "De gekozen map bevat geen HTML-rapporten.")
            return

        try:
            self.server = make_report_server(
                folder,
                on_shutdown=lambda: self.messages.put(("server_stopped", None)),
            )
        except OSError as exc:
            messagebox.showerror("Server niet beschikbaar", str(exc))
            return
        self.server_thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.server_thread.start()
        address = f"http://127.0.0.1:{self.server.server_port}/"
        self.status.set(f"Rapportviewer actief:\n{address}")
        self.start_button.configure(state="disabled")
        self.root.after(100, self._poll_messages)
        if not webbrowser.open(address):
            self.status.set(f"Open dit adres in je browser:\n{address}")

    def _poll_messages(self) -> None:
        try:
            kind, _payload = self.messages.get_nowait()
        except queue.Empty:
            if self.server is not None:
                self.root.after(100, self._poll_messages)
            return
        if kind == "server_stopped":
            self.server = None
            self.server_thread = None
            self.start_button.configure(state="normal")
            self.status.set("De evaluatie is beëindigd.")

    def _stop(self) -> None:
        if self.server is not None:
            self.server.shutdown()
            self.server.server_close()
            self.server = None
        self.server_thread = None
        self.start_button.configure(state="normal")
        self.status.set("De lokale rapportviewer is gestopt.")

    def _close(self) -> None:
        self._stop()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    ReportServerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

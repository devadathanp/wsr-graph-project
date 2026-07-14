#!/usr/bin/env python3
"""
WSR Generator - desktop app.

A one-click front end for non-technical stakeholders: pick the Scrum workbook
(and optionally the Planning workbook) and generate the Weekly Status Report
PowerPoint. No Python knowledge or command line required.
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import threading
import traceback
from datetime import datetime
from pathlib import Path

# Configure matplotlib for headless rendering to a writable cache BEFORE it is
# imported anywhere downstream (charts.py). This keeps the packaged .exe working
# even when the user's home directory is not writable.
os.environ.setdefault("MPLBACKEND", "Agg")
os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "wsr_mplcache"))

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "WSR Generator"
EXCEL_TYPES = [("Excel workbook", "*.xlsm *.xlsx *.xls"), ("All files", "*.*")]


def _reveal(path: Path) -> None:
    """Open a file (or its containing folder) with the OS default handler."""
    try:
        if sys.platform.startswith("darwin"):
            subprocess.run(["open", str(path)], check=False)
        elif os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        pass


class WsrApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(False, False)
        self.configure(padx=18, pady=16)

        self.scrum_var = tk.StringVar()
        self.planning_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d-%m-%Y"))
        self.status_var = tk.StringVar(value="Select your Scrum workbook to begin.")

        self._build_ui()

    def _build_ui(self) -> None:
        header = ttk.Label(self, text="Weekly Status Report Generator", font=("Segoe UI", 15, "bold"))
        header.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        sub = ttk.Label(
            self,
            text="Generate the WSR PowerPoint from your Scrum and Planning Excel files.",
            foreground="#555",
        )
        sub.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 12))

        self._file_row(2, "Scrum workbook *", self.scrum_var, self._pick_scrum)
        self._file_row(3, "Planning workbook", self.planning_var, self._pick_planning)
        self._file_row(4, "Save report as", self.output_var, self._pick_output, save=True)

        ttk.Label(self, text="Report date *").grid(row=5, column=0, sticky="w", pady=4)
        date_entry = ttk.Entry(self, textvariable=self.date_var, width=48)
        date_entry.grid(row=5, column=1, sticky="we", padx=8, pady=4)
        ttk.Label(self, text="dd-mm-yyyy").grid(row=5, column=2, sticky="w", pady=4)

        hint = ttk.Label(
            self,
            text="Used on slides and as the Planned Completion cutoff for pending tables (slides 5–6).",
            foreground="#666",
            wraplength=460,
        )
        hint.grid(row=6, column=0, columnspan=3, sticky="w", pady=(0, 8))

        self.progress = ttk.Progressbar(self, mode="indeterminate", length=420)
        self.progress.grid(row=7, column=0, columnspan=3, sticky="we", pady=(6, 4))

        self.status = ttk.Label(self, textvariable=self.status_var, foreground="#333", wraplength=460)
        self.status.grid(row=8, column=0, columnspan=3, sticky="w")

        self.generate_btn = ttk.Button(self, text="Generate WSR", command=self._on_generate)
        self.generate_btn.grid(row=9, column=0, columnspan=3, sticky="e", pady=(14, 0))

        self.columnconfigure(1, weight=1)

    def _file_row(self, row: int, label: str, var: tk.StringVar, command, save: bool = False) -> None:
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", pady=4)
        entry = ttk.Entry(self, textvariable=var, width=48)
        entry.grid(row=row, column=1, sticky="we", padx=8, pady=4)
        ttk.Button(self, text="Browse…", command=command).grid(row=row, column=2, pady=4)

    def _pick_scrum(self) -> None:
        path = filedialog.askopenfilename(title="Select the Scrum workbook", filetypes=EXCEL_TYPES)
        if not path:
            return
        self.scrum_var.set(path)
        self._autofill_from_scrum(Path(path))

    def _pick_planning(self) -> None:
        path = filedialog.askopenfilename(title="Select the Planning workbook", filetypes=EXCEL_TYPES)
        if path:
            self.planning_var.set(path)

    def _pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save the WSR report",
            defaultextension=".pptx",
            filetypes=[("PowerPoint", "*.pptx")],
            initialfile=self._default_output_name(),
        )
        if path:
            self.output_var.set(path)

    def _autofill_from_scrum(self, scrum: Path) -> None:
        folder = scrum.parent
        if not self.output_var.get():
            self.output_var.set(str(folder / self._default_output_name()))
        if not self.planning_var.get():
            for name in ("Book2.xlsx", "Planning.xlsx", "Planning.xlsm"):
                candidate = folder / name
                if candidate.exists():
                    self.planning_var.set(str(candidate))
                    break
        try:
            from wsr.graph import latest_reported_week

            _, detected_date = latest_reported_week(str(scrum))
            if detected_date:
                self.date_var.set(detected_date)
        except Exception:
            pass
        self.status_var.set("Ready. Confirm the report date, then click 'Generate WSR'.")

    @staticmethod
    def _default_output_name() -> str:
        return f"WSR_Report_{datetime.now():%Y%m%d}.pptx"

    def _on_generate(self) -> None:
        scrum = self.scrum_var.get().strip()
        if not scrum or not Path(scrum).exists():
            messagebox.showerror(APP_TITLE, "Please select a valid Scrum workbook.")
            return
        report_date = self.date_var.get().strip()
        try:
            datetime.strptime(report_date, "%d-%m-%Y")
        except ValueError:
            messagebox.showerror(
                APP_TITLE,
                "Report date must be in dd-mm-yyyy format (e.g. 09-07-2026).",
            )
            return
        output = self.output_var.get().strip() or str(Path(scrum).parent / self._default_output_name())
        self.output_var.set(output)
        planning = self.planning_var.get().strip() or None

        self.generate_btn.config(state="disabled")
        self.progress.start(12)
        self.status_var.set("Generating report… this can take up to a minute.")

        thread = threading.Thread(
            target=self._run_generation,
            args=(scrum, planning, output, report_date),
            daemon=True,
        )
        thread.start()

    def _run_generation(
        self,
        scrum: str,
        planning: str | None,
        output: str,
        report_date: str,
    ) -> None:
        try:
            # Imported here so the window appears instantly and env vars above apply.
            from wsr.report import generate_report

            assets_dir = Path(tempfile.mkdtemp(prefix="wsr_assets_"))
            result = generate_report(
                output_path=output,
                data_file=scrum,
                assets_dir=assets_dir,
                planning_book=planning,
                report_date=report_date,
            )
            self.after(0, self._on_success, Path(result))
        except Exception as exc:  # surfaced to the user in a dialog
            detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
            self.after(0, self._on_failure, str(exc), detail)

    def _on_success(self, result: Path) -> None:
        self.progress.stop()
        self.generate_btn.config(state="normal")
        self.status_var.set(f"Done: {result}")
        if messagebox.askyesno(APP_TITLE, f"Report created:\n{result}\n\nOpen it now?"):
            _reveal(result)

    def _on_failure(self, message: str, detail: str) -> None:
        self.progress.stop()
        self.generate_btn.config(state="normal")
        self.status_var.set("Generation failed.")
        dialog = tk.Toplevel(self)
        dialog.title(f"{APP_TITLE} - Error")
        ttk.Label(dialog, text=message, foreground="#b00020", wraplength=520, padding=12).pack(anchor="w")
        text = tk.Text(dialog, width=80, height=16, wrap="word")
        text.insert("1.0", detail)
        text.config(state="disabled")
        text.pack(padx=12, pady=(0, 12))
        ttk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=(0, 12))


def main() -> None:
    app = WsrApp()
    app.mainloop()


if __name__ == "__main__":
    main()

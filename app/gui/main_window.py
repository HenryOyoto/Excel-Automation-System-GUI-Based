"""CustomTkinter desktop GUI for the Excel Automation System.

The GUI does not do any Excel work itself. It calls:

    app.excel.engine.validate_only  ->  "Validate" button
    app.excel.engine.build_master   ->  "Generate Master Report" button

Everything that touches a workbook lives in app/excel.
"""
from __future__ import annotations

import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from app.excel.engine import build_master, validate_only
from app.models.data_models import ProcessingResult, ReportError
from app.utils import constants as C
from app.utils.file_utils import (
    default_input_dir,
    default_output_dir,
    ensure_dir,
    find_excel_files,
)

# Keep the theme declaration in one place so packaging stays deterministic.
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title(f"{C.APP_NAME} v{C.APP_VERSION}")
        self.geometry("1000x740")
        self.minsize(860, 640)

        # Paths. Default to the project's data\input and data\output folders.
        self.input_dir = ctk.StringVar(value=str(default_input_dir()))
        self.output_dir = ctk.StringVar(value=str(default_output_dir()))

        # Detected / manually added input files, in display order.
        self.selected_files: list[Path] = []
        self._busy = False

        self._build_layout()
        # Auto-detect on startup so the user sees files immediately.
        self.refresh_detected()

    # ------------------------------------------------------------------ layout
    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)  # file list grows
        self.grid_rowconfigure(4, weight=2)  # log grows more

        # Header -----------------------------------------------------------
        header = ctk.CTkFrame(self, corner_radius=0, fg_color=("#1F4E78", "#10344F"))
        header.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            header, text=C.APP_NAME,
            font=ctk.CTkFont(size=24, weight="bold"), text_color="white",
        ).pack(padx=20, pady=(14, 0), anchor="w")
        ctk.CTkLabel(
            header, text=C.APP_TAGLINE,
            font=ctk.CTkFont(size=13), text_color="#D6E4F0",
        ).pack(padx=20, pady=(0, 14), anchor="w")

        # Folder pickers ---------------------------------------------------
        folders = ctk.CTkFrame(self)
        folders.grid(row=1, column=0, sticky="ew", padx=16, pady=(14, 6))
        folders.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(folders, text="Input folder:").grid(
            row=0, column=0, padx=(10, 8), pady=8, sticky="w")
        ctk.CTkEntry(folders, textvariable=self.input_dir).grid(
            row=0, column=1, sticky="ew", pady=8)
        ctk.CTkButton(folders, text="Browse", width=90,
                      command=self.choose_input).grid(row=0, column=2, padx=8, pady=8)

        ctk.CTkLabel(folders, text="Output folder:").grid(
            row=1, column=0, padx=(10, 8), pady=8, sticky="w")
        ctk.CTkEntry(folders, textvariable=self.output_dir).grid(
            row=1, column=1, sticky="ew", pady=8)
        ctk.CTkButton(folders, text="Browse", width=90,
                      command=self.choose_output).grid(row=1, column=2, padx=8, pady=8)

        # Action buttons ---------------------------------------------------
        buttons = ctk.CTkFrame(self, fg_color="transparent")
        buttons.grid(row=2, column=0, sticky="ew", padx=16, pady=6)

        def add_button(col: int, text: str, cmd, primary: bool = False) -> None:
            ctk.CTkButton(
                buttons, text=text, command=cmd, width=155,
                fg_color=("#2E75B6", "#1F4E78") if primary else None,
                hover_color=("#1F4E78", "#10344F") if primary else None,
            ).grid(row=0, column=col, padx=4, pady=4)

        add_button(0, "Scan folder",        self.refresh_detected)
        add_button(1, "Add files…",         self.add_files)
        add_button(2, "Clear list",         self.clear_files)
        add_button(3, "Validate",           self.run_validate)
        add_button(4, "Generate Master Report", self.run_build, primary=True)

        # File list --------------------------------------------------------
        list_frame = ctk.CTkFrame(self)
        list_frame.grid(row=3, column=0, sticky="nsew", padx=16, pady=(6, 0))
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        self.count_label = ctk.CTkLabel(
            list_frame, text="No files detected",
            font=ctk.CTkFont(weight="bold"))
        self.count_label.grid(row=0, column=0, sticky="w", padx=10, pady=6)

        self.file_box = ctk.CTkTextbox(list_frame, wrap="none")
        self.file_box.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.file_box.configure(state="disabled")

        # Log + progress ---------------------------------------------------
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=4, column=0, sticky="nsew", padx=16, pady=10)
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(log_frame, text="Activity log",
                     font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=6)
        self.log_box = ctk.CTkTextbox(log_frame, wrap="word")
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))

        self.progress = ctk.CTkProgressBar(log_frame)
        self.progress.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.progress.set(0)

        self.status = ctk.CTkLabel(self, text="Ready", anchor="w")
        self.status.grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 10))

    # ---------------------------------------------------------------- helpers
    # All widget updates from worker threads go through `after`, which is
    # Tk's only thread-safe entry point.

    def log(self, msg: str) -> None:
        self.after(0, self._log, msg)

    def _log(self, msg: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def set_progress(self, value: float, msg: str = "") -> None:
        def apply() -> None:
            self.progress.set(max(0.0, min(1.0, value)))
            if msg:
                self.status.configure(text=msg)
        self.after(0, apply)

    def _render_files(self) -> None:
        self.file_box.configure(state="normal")
        self.file_box.delete("1.0", "end")
        for p in self.selected_files:
            self.file_box.insert("end", f"{p.name}\n")
        self.file_box.configure(state="disabled")

        n = len(self.selected_files)
        self.count_label.configure(
            text=f"{n} file(s) selected" if n else "No files detected")

    # ------------------------------------------------------------------ paths
    def choose_input(self) -> None:
        folder = filedialog.askdirectory(
            title="Select input folder", initialdir=self.input_dir.get())
        if folder:
            self.input_dir.set(folder)
            self.refresh_detected()

    def choose_output(self) -> None:
        folder = filedialog.askdirectory(
            title="Select output folder", initialdir=self.output_dir.get())
        if folder:
            self.output_dir.set(folder)

    # --------------------------------------------------------------- file list
    def refresh_detected(self) -> None:
        """Scan the input folder and replace the current selection."""
        folder = Path(self.input_dir.get())
        found = find_excel_files(folder)
        self.selected_files = found
        self._render_files()
        self.log(f"Detected {len(found)} Excel file(s) in {folder}")

    def add_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select Excel workbooks",
            filetypes=[("Excel workbooks", "*.xlsx")],
        )
        if not files:
            return
        existing = {p.resolve() for p in self.selected_files}
        added = 0
        for f in files:
            p = Path(f)
            if p.resolve() not in existing:
                self.selected_files.append(p)
                existing.add(p.resolve())
                added += 1
        self._render_files()
        if added:
            self.log(f"Added {added} file(s) manually")

    def clear_files(self) -> None:
        self.selected_files = []
        self._render_files()

    # --------------------------------------------------------------- validate
    def run_validate(self) -> None:
        if self._busy:
            return
        if not self.selected_files:
            messagebox.showwarning(C.APP_NAME, "No files to validate.")
            return
        self._busy = True
        self.set_progress(0.1, "Validating...")
        threading.Thread(target=self._validate_worker, daemon=True).start()

    def _validate_worker(self) -> None:
        try:
            results = validate_only(self.selected_files, log=self.log)
            ok = sum(1 for r in results if r.ok)
            total = len(results)
            self.set_progress(1.0, f"{ok}/{total} files valid")
            self.after(0, lambda: messagebox.showinfo(
                C.APP_NAME,
                f"{ok} of {total} file(s) are valid.\nSee the log for details."))
        except Exception as exc:
            self.log(f"ERROR: {exc}")
            self.set_progress(0.0, "Validation failed")
            self.after(0, lambda: messagebox.showerror(C.APP_NAME, str(exc)))
        finally:
            self._busy = False

    # ------------------------------------------------------------------ build
    def run_build(self) -> None:
        if self._busy:
            return
        if not self.selected_files:
            messagebox.showwarning(C.APP_NAME, "No input files selected.")
            return

        out = Path(self.output_dir.get())
        try:
            ensure_dir(out)
        except OSError as exc:
            messagebox.showerror(C.APP_NAME, f"Cannot use output folder:\n{exc}")
            return

        self._busy = True
        self.set_progress(0.0, "Starting...")
        threading.Thread(target=self._build_worker, args=(out,), daemon=True).start()

    def _build_worker(self, out: Path) -> None:
        try:
            result: ProcessingResult = build_master(
                self.selected_files,
                out,
                rebuild=False,
                progress=self.set_progress,
                log=self.log,
            )
            self.set_progress(1.0, "Complete")
            self.after(0, self._show_success, result)
        except ReportError as exc:
            self.log(f"\nERROR: {exc}")
            self.set_progress(0.0, "Failed")
            self.after(0, lambda: messagebox.showerror(C.APP_NAME, str(exc)))
        except Exception as exc:
            self.log(f"\nUNEXPECTED ERROR: {exc}")
            self.set_progress(0.0, "Failed")
            self.after(0, lambda: messagebox.showerror(
                C.APP_NAME, f"Unexpected error:\n{exc}"))
        finally:
            self._busy = False

    def _show_success(self, r: ProcessingResult) -> None:
        if r.files_processed == 0:
            messagebox.showinfo(C.APP_NAME, r.message)
        else:
            messagebox.showinfo(
                C.APP_NAME,
                f"{r.message}\n\nOutput:\n{r.master_path}")


def run() -> None:
    """Launch the GUI (called from main.py)."""
    App().mainloop()
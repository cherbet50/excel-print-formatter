#!/usr/bin/env python3
"""Desktop app: convert many counter-book Excel files to landscape A4 at once.

Add files (or a whole folder), pick where to save, press Convert.
Works on macOS and Windows. The conversion logic lives in converter.py.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from converter import collect_inputs, convert_many


def open_folder(path: Path) -> None:
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    elif sys.platform.startswith("win"):
        os.startfile(str(path))  # noqa: S606 - opens the user's own output folder
    else:
        subprocess.Popen(["xdg-open", str(path)])


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Print Formatter (A4 landscape)")
        self.geometry("820x600")
        self.minsize(640, 480)
        self.files: list[Path] = []
        self.output_dir = tk.StringVar()
        self.status = tk.StringVar(value="Add .xlsx files to begin.")
        self.events: queue.Queue = queue.Queue()
        self.last_output: Path | None = None
        self._build()

    # ---- layout -----------------------------------------------------------
    def _build(self) -> None:
        pad = {"padx": 10, "pady": 4}

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", **pad)
        ttk.Button(buttons, text="Add files...", command=self.add_files).pack(side="left")
        ttk.Button(buttons, text="Add folder...", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(buttons, text="Remove selected", command=self.remove_selected).pack(side="left")
        ttk.Button(buttons, text="Clear all", command=self.clear_files).pack(side="left", padx=6)

        list_frame = ttk.LabelFrame(self, text="Files to convert")
        list_frame.pack(fill="both", expand=True, **pad)
        scroll = ttk.Scrollbar(list_frame, orient="vertical")
        self.listbox = tk.Listbox(list_frame, selectmode="extended", yscrollcommand=scroll.set)
        scroll.config(command=self.listbox.yview)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        scroll.pack(side="right", fill="y", padx=(0, 6), pady=6)

        out = ttk.Frame(self)
        out.pack(fill="x", **pad)
        ttk.Label(out, text="Save to:").pack(side="left")
        ttk.Entry(out, textvariable=self.output_dir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(out, text="Choose...", command=self.choose_output).pack(side="left")

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)
        self.convert_btn = ttk.Button(bottom, text="Convert all", command=self.start_convert)
        self.convert_btn.pack(side="left")
        self.open_btn = ttk.Button(bottom, text="Open output folder", command=self.open_output,
                                   state="disabled")
        self.open_btn.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=6)

        ttk.Label(self, textvariable=self.status).pack(fill="x", padx=10, pady=(0, 2))
        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.pack(fill="both", expand=False, **pad)
        self.log = tk.Text(log_frame, height=8, wrap="word", state="disabled")
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

    # ---- file list --------------------------------------------------------
    def _add(self, paths) -> None:
        known = set(self.files)
        for path in collect_inputs(paths):
            if path not in known:
                self.files.append(path)
                known.add(path)
                self.listbox.insert("end", path.name)
        if self.files and not self.output_dir.get():
            self.output_dir.set(str(self.files[0].parent / "converted"))
        self.status.set(f"{len(self.files)} file(s) ready.")

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Choose Excel files", filetypes=[("Excel workbook", "*.xlsx")])
        if paths:
            self._add(paths)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Choose a folder with Excel files")
        if folder:
            self._add([folder])

    def remove_selected(self) -> None:
        for index in reversed(self.listbox.curselection()):
            self.listbox.delete(index)
            del self.files[index]
        self.status.set(f"{len(self.files)} file(s) ready.")

    def clear_files(self) -> None:
        self.files.clear()
        self.listbox.delete(0, "end")
        self.status.set("Add .xlsx files to begin.")

    def choose_output(self) -> None:
        folder = filedialog.askdirectory(title="Choose where to save the converted files")
        if folder:
            self.output_dir.set(folder)

    # ---- conversion -------------------------------------------------------
    def start_convert(self) -> None:
        if not self.files:
            messagebox.showinfo("Nothing to convert", "Add at least one .xlsx file first.")
            return
        target = self.output_dir.get().strip()
        if not target:
            messagebox.showinfo("Choose a folder", "Choose where to save the converted files.")
            return
        out = Path(target)
        out.mkdir(parents=True, exist_ok=True)
        self.convert_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.progress.config(maximum=len(self.files), value=0)
        self.last_output = out
        self._write_log(f"Converting {len(self.files)} file(s) to {out}")
        sources = list(self.files)

        def worker() -> None:
            def progress(done, total, name):
                self.events.put(("progress", done, total, name))
            results = convert_many(sources, out, progress)
            self.events.put(("done", results))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll)

    def _poll(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, done, total, name = event
                    self.progress.config(value=done)
                    self.status.set(f"Converted {done}/{total}: {name}")
                elif event[0] == "done":
                    self._finish(event[1])
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _finish(self, results) -> None:
        failed = 0
        for src, report, error in results:
            if error:
                failed += 1
                self._write_log(f"FAILED  {src.name}\n        {error}")
            else:
                self._write_log(
                    f"OK      {src.name}  ->  {Path(report.output).name}  "
                    f"(records {report.records}, pages {report.pages}, scale {report.scale}%)")
                for note in report.warnings:
                    self._write_log(f"        note: {note}")
        ok = len(results) - failed
        self.status.set(f"Finished: {ok} converted, {failed} failed.")
        self.convert_btn.config(state="normal")
        self.open_btn.config(state="normal")
        summary = f"{ok} file(s) converted."
        if failed:
            summary += f"\n{failed} file(s) failed. See the log for details."
        messagebox.showinfo("Conversion finished", summary)

    def open_output(self) -> None:
        if self.last_output and self.last_output.exists():
            open_folder(self.last_output)

    def _write_log(self, text: str) -> None:
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()

#!/usr/bin/env python3
"""
PDS File Scanner – GUI
======================
Graphical front-end for scan_pds.py.
Run this file directly; no command-line arguments needed.
"""

import io
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# Re-use all logic from the existing script
from scan_pds import main as run_scan


# ── Redirect stdout into the GUI log ─────────────────────────────────────────

class _QueueWriter(io.TextIOBase):
    """Writes print() output to a thread-safe queue."""
    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, s: str) -> int:
        if s:
            self._q.put(s)
        return len(s)

    def flush(self):
        pass


# ── Main window ───────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDS File Scanner")
        self.resizable(True, True)
        self.minsize(640, 480)

        self._log_queue: queue.Queue = queue.Queue()
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # ── Input frame ──────────────────────────────────────────────────────
        frm = ttk.LabelFrame(self, text="Inputs", padding=10)
        frm.pack(fill="x", **pad)
        frm.columnconfigure(1, weight=1)

        # PDS folder
        ttk.Label(frm, text="PDS Folder:").grid(row=0, column=0, sticky="w", pady=4)
        self._pds_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self._pds_var).grid(
            row=0, column=1, sticky="ew", padx=(6, 4))
        ttk.Button(frm, text="Browse…", command=self._browse_folder).grid(
            row=0, column=2)

        # Tracker file
        ttk.Label(frm, text="Tracker File:").grid(row=1, column=0, sticky="w", pady=4)
        self._tracker_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self._tracker_var).grid(
            row=1, column=1, sticky="ew", padx=(6, 4))
        ttk.Button(frm, text="Browse…", command=self._browse_file).grid(
            row=1, column=2)

        # ── Run button ────────────────────────────────────────────────────────
        self._run_btn = ttk.Button(
            self, text="▶  Run Scan", command=self._start_scan)
        self._run_btn.pack(pady=(0, 4))

        # ── Progress bar ──────────────────────────────────────────────────────
        self._progress = ttk.Progressbar(self, mode="indeterminate")
        self._progress.pack(fill="x", padx=10, pady=(0, 4))

        # ── Log area ──────────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self, text="Log", padding=6)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._log = scrolledtext.ScrolledText(
            log_frame, state="disabled", wrap="word",
            font=("Consolas", 9), background="#1e1e1e", foreground="#d4d4d4",
            insertbackground="white")
        self._log.pack(fill="both", expand=True)

        # colour tags
        self._log.tag_config("ok",    foreground="#4ec9b0")
        self._log.tag_config("skip",  foreground="#dcdcaa")
        self._log.tag_config("dup",   foreground="#9cdcfe")
        self._log.tag_config("err",   foreground="#f44747")
        self._log.tag_config("head",  foreground="#569cd6")

    # ── Browse helpers ────────────────────────────────────────────────────────

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select PDS Folder")
        if path:
            self._pds_var.set(path)

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Tracker File",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")])
        if path:
            self._tracker_var.set(path)

    # ── Scan execution ────────────────────────────────────────────────────────

    def _start_scan(self):
        pds_folder = self._pds_var.get().strip()
        tracker_file = self._tracker_var.get().strip()

        if not pds_folder:
            messagebox.showwarning("Missing input", "Please select a PDS folder.")
            return
        if not tracker_file:
            messagebox.showwarning("Missing input", "Please select a tracker file.")
            return

        self._run_btn.config(state="disabled")
        self._progress.start(10)
        self._clear_log()
        self._log_line("Starting scan…\n", "head")

        threading.Thread(
            target=self._run_scan_thread,
            args=(pds_folder, tracker_file),
            daemon=True,
        ).start()

        self.after(50, self._poll_log)

    def _run_scan_thread(self, pds_folder: str, tracker_file: str):
        import sys
        writer = _QueueWriter(self._log_queue)
        old_stdout = sys.stdout
        sys.stdout = writer
        try:
            run_scan(pds_folder, tracker_file)
        except Exception as exc:
            self._log_queue.put(f"\nFATAL ERROR: {exc}\n")
        finally:
            sys.stdout = old_stdout
            self._log_queue.put(None)   # sentinel → scan finished

    # ── Log polling (runs on main thread via after()) ─────────────────────────

    def _poll_log(self):
        try:
            while True:
                item = self._log_queue.get_nowait()
                if item is None:
                    self._on_scan_done()
                    return
                self._append_log(item)
        except queue.Empty:
            pass
        self.after(50, self._poll_log)

    def _on_scan_done(self):
        self._progress.stop()
        self._run_btn.config(state="normal")
        self._log_line("\nScan complete.\n", "head")

    # ── Log helpers ───────────────────────────────────────────────────────────

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _append_log(self, text: str):
        tag = ""
        upper = text.upper()
        if "[OK]" in upper:
            tag = "ok"
        elif "[SKIP]" in upper:
            tag = "skip"
        elif "[DUP]" in upper:
            tag = "dup"
        elif "[ERR]" in upper or "ERROR" in upper or "FATAL" in upper:
            tag = "err"
        elif text.startswith("=") or text.startswith("Output") or text.startswith("Loading"):
            tag = "head"
        self._log_line(text, tag)

    def _log_line(self, text: str, tag: str = ""):
        self._log.config(state="normal")
        if tag:
            self._log.insert("end", text, tag)
        else:
            self._log.insert("end", text)
        self._log.see("end")
        self._log.config(state="disabled")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()

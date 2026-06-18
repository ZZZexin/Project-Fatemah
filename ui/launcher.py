"""TV Pipeline Launcher — single page: organise → select targets → batch-convert.

Run from project root:
    python -m ui.launcher
    python ui/launcher.py
"""

from __future__ import annotations

import csv
import json
import logging
import queue
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ── path setup ──────────────────────────────────────────────────────────────
_UI_DIR = Path(__file__).resolve().parent
_ROOT = _UI_DIR.parent
for _p in (str(_ROOT), str(_UI_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

# ── project imports ─────────────────────────────────────────────────────────
try:
    from organize_hole_files import build_plan, apply_plan, extract_zips
except ImportError:
    from ui.organize_hole_files import build_plan, apply_plan, extract_zips

try:
    from hed_target_selector import (
        find_hed_runs, HedRun, SELECTABLE_TYPES,
        default_run_label, ALL_CHOICE, ScrollableFrame,
    )
except ImportError:
    from ui.hed_target_selector import (
        find_hed_runs, HedRun, SELECTABLE_TYPES,
        default_run_label, ALL_CHOICE, ScrollableFrame,
    )

DEFAULT_TARGETS_JSON = _ROOT / "config" / "selected_hed_targets.json"
_STATE_FILE = _ROOT / "config" / "launcher_state.json"


def _load_state() -> dict:
    try:
        return json.loads(_STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_state(key: str, value: str) -> None:
    state = _load_state()
    state[key] = value
    _STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    _STATE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")


# ── queue logging handler ───────────────────────────────────────────────────

class _QueueHandler(logging.Handler):
    def __init__(self, q: queue.Queue) -> None:
        super().__init__()
        self.q = q

    def emit(self, record: logging.LogRecord) -> None:
        self.q.put(self.format(record))


# ═══════════════════════════════════════════════════════════════════════════
# Section 1 — Organise
# ═══════════════════════════════════════════════════════════════════════════

class OrganiseSection(ttk.LabelFrame):
    def __init__(self, parent, on_organized=None) -> None:
        super().__init__(parent, text=" 1. Organise Files ", padding=8)
        self.on_organized = on_organized
        self._actions: list = []
        state = _load_state()
        self._source_var = tk.StringVar(value=state.get("last_source_dir", ""))
        self._output_var = tk.StringVar(value=state.get("last_output_dir", ""))
        self._mode_var = tk.StringVar(value="move")
        self._zip_var = tk.BooleanVar(value=True)
        self._status_var = tk.StringVar(value="Select a source folder and click Scan.")
        self._build_ui()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Source:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Entry(top, textvariable=self._source_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(top, text="Browse", command=self._browse_source).grid(row=0, column=2, padx=(6, 0))

        ttk.Label(top, text="Output:").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=(4, 0))
        ttk.Entry(top, textvariable=self._output_var).grid(row=1, column=1, sticky="ew", pady=(4, 0))
        ttk.Button(top, text="Browse", command=self._browse_output).grid(row=1, column=2, padx=(6, 0), pady=(4, 0))
        ttk.Label(top, text="(blank = same as source)").grid(row=1, column=3, sticky="w", padx=(6, 0))

        opts = ttk.Frame(self)
        opts.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        ttk.Checkbutton(opts, text="Extract ZIPs (recursive)", variable=self._zip_var).pack(side="left")
        ttk.Label(opts, text="Mode:").pack(side="left", padx=(16, 4))
        ttk.Combobox(opts, textvariable=self._mode_var, values=["move", "copy"],
                     state="readonly", width=6).pack(side="left")
        ttk.Label(opts, text="Structure: HoleID / OTV · ATV · GPX · Reports").pack(side="left", padx=(16, 0))
        ttk.Button(opts, text="Scan", command=self._scan).pack(side="left", padx=(20, 0))
        ttk.Button(opts, text="Apply  ↓ Targets", command=self._apply).pack(side="left", padx=(8, 0))
        ttk.Label(opts, textvariable=self._status_var).pack(side="left", padx=(12, 0))

        frame = ttk.Frame(self)
        frame.grid(row=2, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        cols = ("hole", "type", "source", "destination")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=7)
        for col, text, width, stretch in [
            ("hole", "Hole", 130, False),
            ("type", "Type", 65, False),
            ("source", "Source", 380, True),
            ("destination", "Destination", 380, True),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, stretch=stretch)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def _browse_source(self) -> None:
        d = filedialog.askdirectory(initialdir=self._source_var.get() or str(Path.home()))
        if d:
            self._source_var.set(d)
            _save_state("last_source_dir", d)

    def _browse_output(self) -> None:
        d = filedialog.askdirectory(
            initialdir=self._output_var.get() or self._source_var.get() or str(Path.home())
        )
        if d:
            self._output_var.set(d)
            _save_state("last_output_dir", d)

    def _resolve_dirs(self):
        src = self._source_var.get().strip()
        if not src:
            messagebox.showerror("No source", "Select a source folder first.")
            return None
        source_dir = Path(src).expanduser().resolve()
        if not source_dir.is_dir():
            messagebox.showerror("Not found", f"Folder not found:\n{source_dir}")
            return None
        out = self._output_var.get().strip()
        output_dir = Path(out).expanduser().resolve() if out else source_dir
        return source_dir, output_dir

    def _scan(self) -> None:
        dirs = self._resolve_dirs()
        if not dirs:
            return
        source_dir, output_dir = dirs
        self._status_var.set("Scanning…")
        self.update()
        self._actions = build_plan(source_dir, output_dir, extract_zip=self._zip_var.get())
        self.tree.delete(*self.tree.get_children())
        for a in self._actions:
            self.tree.insert("", "end", values=(a.hole, a.data_type, str(a.source), str(a.destination)))
        holes = len({a.hole for a in self._actions})
        self._status_var.set(
            f"{len(self._actions)} files across {holes} holes planned." if self._actions
            else "Nothing to sort — files already in place."
        )

    def _apply(self) -> None:
        dirs = self._resolve_dirs()
        if not dirs:
            return
        _, output_dir = dirs
        if not self._actions:
            self._status_var.set("Nothing to sort — proceeding.")
            if self.on_organized:
                self.on_organized(output_dir)
            return
        try:
            apply_plan(self._actions, mode=self._mode_var.get())
        except OSError as exc:
            messagebox.showerror("Apply failed", str(exc))
            return
        verb = "moved" if self._mode_var.get() == "move" else "copied"
        self._status_var.set(f"Done — {len(self._actions)} files {verb}.")
        self._actions = []
        self.tree.delete(*self.tree.get_children())
        if self.on_organized:
            self.on_organized(output_dir)


# ═══════════════════════════════════════════════════════════════════════════
# Section 2 — Select Targets
# ═══════════════════════════════════════════════════════════════════════════

class TargetsSection(ttk.LabelFrame):
    def __init__(self, parent, on_saved=None) -> None:
        super().__init__(parent, text=" 2. Select Targets ", padding=8)
        self.on_saved = on_saved
        self._search_var = tk.StringVar(value=_load_state().get("last_targets_dir", ""))
        self._status_var = tk.StringVar(value="Select a folder and scan for .hed files.")
        self.run_lookup: dict[str, dict[str, list[HedRun]]] = {}
        self.include_holes: dict[str, tk.BooleanVar] = {}
        self.selected_labels: dict[tuple[str, str], tk.StringVar] = {}
        self.path_labels: dict[tuple[str, str], tk.StringVar] = {}
        self._build_ui()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Search folder:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Entry(top, textvariable=self._search_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(top, text="Browse", command=self._browse).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(top, text="Scan", command=self._scan).grid(row=0, column=3, padx=(6, 0))

        acts = ttk.Frame(top)
        acts.grid(row=1, column=0, columnspan=4, pady=(6, 0), sticky="ew")
        ttk.Button(acts, text="Select all", command=self._select_all).pack(side="left")
        ttk.Button(acts, text="Deselect all", command=self._deselect_all).pack(side="left", padx=(6, 0))
        ttk.Button(acts, text="Save  ↓ Convert", command=self._save).pack(side="left", padx=(6, 0))
        ttk.Label(acts, textvariable=self._status_var).pack(side="left", padx=(12, 0))

        self.table = ScrollableFrame(self)
        self.table.grid(row=1, column=0, sticky="nsew", pady=(4, 0))

    def load_from_dir(self, path: Path) -> None:
        """Auto-called after Organise applies — scans the output directory."""
        self._search_var.set(str(path))
        _save_state("last_targets_dir", str(path))
        self._scan()

    def _browse(self) -> None:
        d = filedialog.askdirectory(initialdir=self._search_var.get() or str(Path.home()))
        if d:
            self._search_var.set(d)
            _save_state("last_targets_dir", d)
            self._scan()

    def _scan(self) -> None:
        root = Path(self._search_var.get()).expanduser()
        if not root.is_dir():
            messagebox.showerror("Not found", f"Folder not found:\n{root}")
            return
        self.run_lookup = find_hed_runs(root)
        self._render_rows()
        total_runs = sum(len(runs) for tr in self.run_lookup.values() for runs in tr.values())
        self._status_var.set(
            f"Found {total_runs} .hed files across {len(self.run_lookup)} holes."
        )

    def _render_rows(self) -> None:
        for child in self.table.inner.winfo_children():
            child.destroy()
        self.include_holes.clear()
        self.selected_labels.clear()
        self.path_labels.clear()

        headers = ("Include", "Hole", "OPTV / OTV / OBI", "BHTV / ATV / ABI")
        for col, text in enumerate(headers):
            ttk.Label(self.table.inner, text=text, font=("", 10, "bold")).grid(
                row=0, column=col, sticky="ew", padx=6, pady=(0, 6)
            )
        self.table.inner.columnconfigure(0, weight=0, minsize=70)
        self.table.inner.columnconfigure(1, weight=0, minsize=140)
        self.table.inner.columnconfigure(2, weight=1, minsize=280)
        self.table.inner.columnconfigure(3, weight=1, minsize=280)

        for row_i, (hole, type_runs) in enumerate(self.run_lookup.items(), start=1):
            inc = tk.BooleanVar(value=True)
            self.include_holes[hole] = inc
            ttk.Checkbutton(self.table.inner, variable=inc).grid(
                row=row_i, column=0, sticky="w", padx=6, pady=4
            )
            ttk.Label(self.table.inner, text=hole).grid(
                row=row_i, column=1, sticky="w", padx=6, pady=4
            )
            self._render_selector(row_i, hole, "OPTV", type_runs.get("OPTV", []), 2)
            self._render_selector(row_i, hole, "BHTV", type_runs.get("BHTV", []), 3)

    def _render_selector(self, row_i, hole, data_type, runs, col) -> None:
        choices = [r.label for r in runs] + ([ALL_CHOICE] if runs else [])
        sel = tk.StringVar(value=default_run_label(data_type, runs))
        sel_path = tk.StringVar(value="")
        key = (hole, data_type)
        self.selected_labels[key] = sel
        self.path_labels[key] = sel_path
        if runs:
            self.after_idle(lambda h=hole, t=data_type: self._update_path(h, t))
        combo = ttk.Combobox(
            self.table.inner, textvariable=sel, values=choices,
            state="readonly" if runs else "disabled", width=36,
        )
        combo.grid(row=row_i, column=col, sticky="ew", padx=6, pady=4)
        combo.bind("<<ComboboxSelected>>", lambda _e, h=hole, t=data_type: self._update_path(h, t))

    def _update_path(self, hole, data_type) -> None:
        key = (hole, data_type)
        sel = self.selected_labels[key].get()
        if sel == ALL_CHOICE:
            self.path_labels[key].set(ALL_CHOICE)
            return
        for run in self.run_lookup.get(hole, {}).get(data_type, []):
            if run.label == sel:
                self.path_labels[key].set(str(run.path))
                return
        self.path_labels[key].set("")

    def _selected_runs_for(self, hole, data_type) -> list[HedRun]:
        key = (hole, data_type)
        sel = self.selected_labels[key].get()
        runs = self.run_lookup.get(hole, {}).get(data_type, [])
        return runs if sel == ALL_CHOICE else [r for r in runs if r.label == sel]

    def _select_all(self) -> None:
        for v in self.include_holes.values():
            v.set(True)

    def _deselect_all(self) -> None:
        for v in self.include_holes.values():
            v.set(False)

    def get_targets(self) -> list[dict]:
        targets = []
        for hole in sorted(self.run_lookup):
            inc = self.include_holes.get(hole)
            if not inc or not inc.get():
                continue
            for data_type in SELECTABLE_TYPES:
                self._update_path(hole, data_type)
                for run in self._selected_runs_for(hole, data_type):
                    targets.append({
                        "hole": hole,
                        "data_type": data_type,
                        "run": run.label,
                        "parent_directory": str(run.path.parent),
                        "file_name": run.path.name,
                        "path": str(run.path),
                    })
        return targets

    def _save(self) -> None:
        targets = self.get_targets()
        if not targets:
            messagebox.showwarning("Nothing to save", "Scan first and select at least one run.")
            return
        DEFAULT_TARGETS_JSON.parent.mkdir(parents=True, exist_ok=True)
        DEFAULT_TARGETS_JSON.write_text(json.dumps(targets, indent=2), encoding="utf-8")
        csv_path = DEFAULT_TARGETS_JSON.with_suffix(".csv")
        with csv_path.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(
                f, fieldnames=["hole", "data_type", "run", "parent_directory", "file_name", "path"]
            )
            w.writeheader()
            w.writerows(targets)
        self._status_var.set(f"Saved {len(targets)} targets.")
        if self.on_saved:
            self.on_saved(DEFAULT_TARGETS_JSON)


# ═══════════════════════════════════════════════════════════════════════════
# Section 3 — Convert
# ═══════════════════════════════════════════════════════════════════════════

class ConvertSection(ttk.LabelFrame):
    def __init__(self, parent) -> None:
        super().__init__(parent, text=" 3. Batch Convert ", padding=8)
        self._targets_path = DEFAULT_TARGETS_JSON
        self._summary_var = tk.StringVar(value="No targets loaded.")
        self._running = False
        self._stop_event = threading.Event()
        self._log_q: queue.Queue = queue.Queue()
        self._log_handler: logging.Handler | None = None
        self._build_ui()
        if DEFAULT_TARGETS_JSON.exists():
            self._load_targets(DEFAULT_TARGETS_JSON)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 2))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Targets JSON:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._path_var = tk.StringVar(value=str(self._targets_path))
        ttk.Entry(top, textvariable=self._path_var, state="readonly").grid(row=0, column=1, sticky="ew")
        ttk.Button(top, text="Browse", command=self._browse_targets).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(
            top, text="Reload",
            command=lambda: self._load_targets(Path(self._path_var.get())),
        ).grid(row=0, column=3, padx=(6, 0))

        ttk.Label(self, textvariable=self._summary_var, foreground="gray").grid(
            row=1, column=0, sticky="w", pady=(0, 4)
        )

        ctrl = ttk.Frame(self)
        ctrl.grid(row=2, column=0, sticky="w", pady=(0, 6))
        self._run_btn = ttk.Button(ctrl, text="Run Convert", command=self._start)
        self._run_btn.pack(side="left")
        self._stop_btn = ttk.Button(ctrl, text="Stop", command=self._stop, state="disabled")
        self._stop_btn.pack(side="left", padx=(6, 0))
        ttk.Button(ctrl, text="Clear log", command=self._clear_log).pack(side="left", padx=(6, 0))

        log_frame = ttk.LabelFrame(self, text="Convert log", padding=4)
        log_frame.grid(row=3, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap="word", state="disabled", font=("Consolas", 9), height=10
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.tag_config("ERR", foreground="#cc0000")
        self.log_text.tag_config("WARN", foreground="#cc6600")
        self.log_text.tag_config("OK", foreground="#007700")

    def load_targets(self, json_path: Path) -> None:
        """Auto-called after Targets saves."""
        self._load_targets(json_path)

    def _browse_targets(self) -> None:
        f = filedialog.askopenfilename(
            initialdir=str(DEFAULT_TARGETS_JSON.parent),
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if f:
            self._load_targets(Path(f))

    def _load_targets(self, json_path: Path) -> None:
        self._targets_path = json_path
        self._path_var.set(str(json_path))
        if not json_path.exists():
            self._summary_var.set("File not found.")
            return
        try:
            targets = json.loads(json_path.read_text(encoding="utf-8"))
            optv = sum(1 for t in targets if t.get("data_type") == "OPTV")
            bhtv = sum(1 for t in targets if t.get("data_type") == "BHTV")
            self._summary_var.set(
                f"Loaded {len(targets)} targets — {optv} OPTV, {bhtv} BHTV"
            )
        except Exception as exc:
            self._summary_var.set(f"Load error: {exc}")

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        upper = text.upper()
        tag = ""
        if "ERROR" in upper or "FAIL" in upper:
            tag = "ERR"
        elif "WARNING" in upper or "WARN" in upper:
            tag = "WARN"
        elif "OK" in upper or " DONE" in upper:
            tag = "OK"
        self.log_text.insert("end", text + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _poll_log(self) -> None:
        try:
            while True:
                msg = self._log_q.get_nowait()
                if msg is None:
                    self._on_done()
                    return
                self._append_log(msg)
        except queue.Empty:
            pass
        if self._running:
            self.after(100, self._poll_log)

    def _start(self) -> None:
        if self._running:
            return
        if not self._targets_path.exists():
            messagebox.showerror("No targets", f"File not found:\n{self._targets_path}")
            return
        self._running = True
        self._stop_event.clear()
        self._run_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._append_log(f"=== Starting convert: {self._targets_path} ===")

        root_log = logging.getLogger()
        root_log.setLevel(logging.INFO)
        handler = _QueueHandler(self._log_q)
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)-8s %(message)s"))
        root_log.addHandler(handler)
        self._log_handler = handler
        self._root_log = root_log

        targets_path = str(self._targets_path)
        stop_event = self._stop_event

        def worker() -> None:
            try:
                from modules.convert.batch_convert import run_all
                run_all(targets_path, stop_event=stop_event)
            except Exception as exc:
                logging.error("Convert failed: %s", exc, exc_info=True)
            finally:
                self._log_q.put(None)

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_log)

    def _stop(self) -> None:
        self._stop_event.set()
        self._append_log("=== Stop requested — finishing current file then stopping ===")
        self._run_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")

    def _on_done(self) -> None:
        self._running = False
        self._run_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")
        if self._log_handler and hasattr(self, "_root_log"):
            self._root_log.removeHandler(self._log_handler)
            self._log_handler = None
        self._append_log("=== Convert complete ===")


# ═══════════════════════════════════════════════════════════════════════════
# Main application
# ═══════════════════════════════════════════════════════════════════════════

class LauncherApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("TV Pipeline Launcher")
        self.geometry("1200x960")
        self.minsize(900, 700)

        pw = ttk.PanedWindow(self, orient="vertical")
        pw.pack(fill="both", expand=True, padx=8, pady=8)

        # Build back-to-front so callbacks can reference later sections
        convert_sec = ConvertSection(pw)
        targets_sec = TargetsSection(
            pw,
            on_saved=lambda p: convert_sec.load_targets(p),
        )
        organise_sec = OrganiseSection(
            pw,
            on_organized=lambda p: targets_sec.load_from_dir(p),
        )

        pw.add(organise_sec, weight=2)
        pw.add(targets_sec, weight=3)
        pw.add(convert_sec, weight=2)


def main() -> None:
    app = LauncherApp()
    app.mainloop()


if __name__ == "__main__":
    main()

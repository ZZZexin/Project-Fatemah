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
# When frozen by PyInstaller, keep config/logs next to the .exe, not in the
# temp extraction dir (_MEIPASS).
if getattr(sys, "frozen", False):
    _ROOT = Path(sys.executable).resolve().parent
else:
    _ROOT = _UI_DIR.parent
    for _p in (str(_ROOT), str(_UI_DIR)):
        if _p not in sys.path:
            sys.path.insert(0, _p)

# ── project imports ─────────────────────────────────────────────────────────
try:
    from organize_hole_files import (build_plan, apply_plan, extract_zips,
                                      delete_unwanted_files, delete_empty_dirs,
                                      discover_holes, ensure_hole_structure,
                                      archive_zips)
except ImportError:
    from ui.organize_hole_files import (build_plan, apply_plan, extract_zips,
                                        delete_unwanted_files, delete_empty_dirs,
                                        discover_holes, ensure_hole_structure,
                                        archive_zips)

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


# ── shared folder-row helper ────────────────────────────────────────────────

def _folder_row(parent, label: str, var: tk.StringVar, browse_cmd, hint: str = "") -> None:
    """One labelled folder entry + Browse button, spanning a grid row."""
    ttk.Label(parent, text=label, width=8, anchor="w").grid(
        row=parent._next_row, column=0, sticky="w", padx=(0, 6), pady=3
    )
    ttk.Entry(parent, textvariable=var).grid(
        row=parent._next_row, column=1, sticky="ew", pady=3
    )
    ttk.Button(parent, text="Browse…", command=browse_cmd, width=9).grid(
        row=parent._next_row, column=2, sticky="e", padx=(6, 0), pady=3
    )
    if hint:
        ttk.Label(parent, text=hint, foreground="gray").grid(
            row=parent._next_row, column=3, sticky="w", padx=(8, 0), pady=3
        )
    parent._next_row += 1


# ═══════════════════════════════════════════════════════════════════════════
# Section 1 — Organise
# ═══════════════════════════════════════════════════════════════════════════

class OrganiseSection(ttk.LabelFrame):
    def __init__(self, parent, on_organized=None) -> None:
        super().__init__(parent, text="  Step 1 — Organise Files  ", padding=10)
        self.on_organized = on_organized
        self._actions: list = []
        state = _load_state()
        self._source_var = tk.StringVar(value=state.get("last_source_dir", ""))
        self._output_var = tk.StringVar(value=state.get("last_output_dir", ""))
        self._mode_var = tk.StringVar(value="move")
        self._zip_var = tk.BooleanVar(value=True)
        self._status_var = tk.StringVar(value="Select a source folder and scan.")
        self._build_ui()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # ── folder rows ──────────────────────────────────────────────────
        folders = ttk.Frame(self)
        folders.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        folders.columnconfigure(1, weight=1)
        folders._next_row = 0
        _folder_row(folders, "Source:", self._source_var, self._browse_source)
        _folder_row(folders, "Output:", self._output_var, self._browse_output,
                    hint="leave blank to organise in-place")

        # ── options + action bar ─────────────────────────────────────────
        bar = ttk.Frame(self)
        bar.grid(row=1, column=0, sticky="ew", pady=(0, 6))

        ttk.Label(bar, text="Mode:").pack(side="left")
        ttk.Radiobutton(bar, text="Move", variable=self._mode_var, value="move").pack(
            side="left", padx=(6, 0))
        ttk.Radiobutton(bar, text="Copy", variable=self._mode_var, value="copy").pack(
            side="left", padx=(4, 0))

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=12, pady=2)

        ttk.Checkbutton(bar, text="Extract ZIPs recursively",
                        variable=self._zip_var).pack(side="left")

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=12, pady=2)

        ttk.Button(bar, text="Scan", command=self._scan, width=8).pack(side="left")
        ttk.Button(bar, text="Organise  ▶", command=self._apply).pack(
            side="left", padx=(8, 0))

        ttk.Label(bar, textvariable=self._status_var,
                  foreground="gray").pack(side="left", padx=(14, 0))

        # ── preview table ─────────────────────────────────────────────────
        frame = ttk.Frame(self)
        frame.grid(row=2, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        cols = ("hole", "type", "source", "destination")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=7)
        for col, text, width, stretch in [
            ("hole",        "Hole ID",    130, False),
            ("type",        "Type",        60, False),
            ("source",      "Source",     360, True),
            ("destination", "Destination",360, True),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, stretch=stretch, minwidth=50)

        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        ttk.Label(self, text="Sorted into: HoleID / OTV  ·  ATV  ·  GPX   |   PDFs and LAS files removed after move",
                  foreground="gray").grid(row=3, column=0, sticky="w", pady=(4, 0))

    # ── callbacks ──────────────────────────────────────────────────────────

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
            self.tree.insert("", "end",
                             values=(a.hole, a.data_type, str(a.source), str(a.destination)))
        holes = len({a.hole for a in self._actions})
        self._status_var.set(
            f"{len(self._actions)} files across {holes} holes planned." if self._actions
            else "Nothing to move — files already in place."
        )

    def _apply(self) -> None:
        dirs = self._resolve_dirs()
        if not dirs:
            return
        source_dir, output_dir = dirs
        # Re-derive from the current disk state so a stale scan can't trigger
        # work — and so an already-sorted folder does nothing.
        if not self._actions:
            self._actions = build_plan(source_dir, output_dir,
                                       extract_zip=self._zip_var.get())
        if not self._actions:
            self._status_var.set("Already organised — nothing to do.")
            return
        holes = {a.hole for a in self._actions}
        try:
            apply_plan(self._actions, mode=self._mode_var.get())
        except OSError as exc:
            messagebox.showerror("Apply failed", str(exc))
            return
        move_mode = self._mode_var.get() == "move"
        verb = "moved" if move_mode else "copied"
        deleted = delete_unwanted_files(source_dir) if move_mode else 0
        # In move mode, tuck the processed zips into sorted/ so the next run
        # finds nothing to do, then sweep empty dirs.
        archived = archive_zips(source_dir) if move_mode else 0
        removed = delete_empty_dirs(source_dir) if move_mode else 0
        # (re)create the standard OTV/ATV/GPX layout for every hole — both the
        # ones we just moved and any already present.
        ensure_hole_structure(output_dir, holes | discover_holes(output_dir))
        parts = [f"{len(self._actions)} files {verb}"]
        if deleted:
            parts.append(f"{deleted} PDF/LAS deleted")
        if archived:
            parts.append(f"{archived} zips → sorted/")
        if removed:
            parts.append(f"{removed} empty dirs removed")
        self._status_var.set("Done — " + ", ".join(parts) + ".")
        self._actions = []
        self.tree.delete(*self.tree.get_children())
        if self.on_organized:
            self.on_organized(output_dir)


# ═══════════════════════════════════════════════════════════════════════════
# Section 2 — Select Targets
# ═══════════════════════════════════════════════════════════════════════════

class TargetsSection(ttk.LabelFrame):
    def __init__(self, parent, on_saved=None) -> None:
        super().__init__(parent, text="  Step 2 — Select Targets  ", padding=10)
        self.on_saved = on_saved
        self._search_var = tk.StringVar(value=_load_state().get("last_targets_dir", ""))
        self._status_var = tk.StringVar(value="Browse to a folder or let Step 1 populate this automatically.")
        self.run_lookup: dict[str, dict[str, list[HedRun]]] = {}
        self.include_holes: dict[str, tk.BooleanVar] = {}
        self.selected_labels: dict[tuple[str, str], tk.StringVar] = {}
        self.path_labels: dict[tuple[str, str], tk.StringVar] = {}
        self._build_ui()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # ── top bar ───────────────────────────────────────────────────────
        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Folder:", width=8, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Entry(top, textvariable=self._search_var).grid(
            row=0, column=1, sticky="ew")
        ttk.Button(top, text="Browse…", command=self._browse, width=9).grid(
            row=0, column=2, sticky="e", padx=(6, 0))
        ttk.Button(top, text="Scan", command=self._scan, width=8).grid(
            row=0, column=3, padx=(6, 0))

        # ── action bar ────────────────────────────────────────────────────
        acts = ttk.Frame(top)
        acts.grid(row=1, column=0, columnspan=4, pady=(8, 0), sticky="ew")
        ttk.Button(acts, text="Select all",   command=self._select_all).pack(side="left")
        ttk.Button(acts, text="Deselect all", command=self._deselect_all).pack(
            side="left", padx=(6, 0))
        ttk.Separator(acts, orient="vertical").pack(side="left", fill="y", padx=12, pady=2)
        ttk.Button(acts, text="Save  ▶  Load Convert", command=self._save).pack(side="left")
        ttk.Label(acts, textvariable=self._status_var,
                  foreground="gray").pack(side="left", padx=(14, 0))

        # ── scrollable hole table ─────────────────────────────────────────
        self.table = ScrollableFrame(self)
        self.table.grid(row=1, column=0, sticky="nsew", pady=(0, 0))

    # ── callbacks ──────────────────────────────────────────────────────────

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
            f"{total_runs} .hed files across {len(self.run_lookup)} holes found."
        )

    def _render_rows(self) -> None:
        for child in self.table.inner.winfo_children():
            child.destroy()
        self.include_holes.clear()
        self.selected_labels.clear()
        self.path_labels.clear()

        headers = ("Include", "Hole ID", "OPTV / OTV / OBI", "BHTV / ATV / ABI")
        for col, text in enumerate(headers):
            ttk.Label(self.table.inner, text=text,
                      font=("", 10, "bold")).grid(
                row=0, column=col, sticky="ew", padx=6, pady=(0, 6))
        ttk.Separator(self.table.inner, orient="horizontal").grid(
            row=1, column=0, columnspan=4, sticky="ew", pady=(0, 4))

        self.table.inner.columnconfigure(0, weight=0, minsize=70)
        self.table.inner.columnconfigure(1, weight=0, minsize=140)
        self.table.inner.columnconfigure(2, weight=1, minsize=280)
        self.table.inner.columnconfigure(3, weight=1, minsize=280)

        for row_i, (hole, type_runs) in enumerate(self.run_lookup.items(), start=2):
            inc = tk.BooleanVar(value=True)
            self.include_holes[hole] = inc
            ttk.Checkbutton(self.table.inner, variable=inc).grid(
                row=row_i, column=0, sticky="w", padx=6, pady=3)
            ttk.Label(self.table.inner, text=hole).grid(
                row=row_i, column=1, sticky="w", padx=6, pady=3)
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
        combo.grid(row=row_i, column=col, sticky="ew", padx=6, pady=3)
        combo.bind("<<ComboboxSelected>>",
                   lambda _e, h=hole, t=data_type: self._update_path(h, t))

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
            messagebox.showwarning("Nothing to save",
                                   "Scan first and select at least one run.")
            return
        DEFAULT_TARGETS_JSON.parent.mkdir(parents=True, exist_ok=True)
        DEFAULT_TARGETS_JSON.write_text(
            json.dumps(targets, indent=2), encoding="utf-8")
        csv_path = DEFAULT_TARGETS_JSON.with_suffix(".csv")
        with csv_path.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(
                f, fieldnames=["hole", "data_type", "run",
                               "parent_directory", "file_name", "path"])
            w.writeheader()
            w.writerows(targets)
        optv = sum(1 for t in targets if t["data_type"] == "OPTV")
        bhtv = sum(1 for t in targets if t["data_type"] == "BHTV")
        self._status_var.set(f"Saved — {optv} OPTV, {bhtv} BHTV targets.")
        if self.on_saved:
            self.on_saved(DEFAULT_TARGETS_JSON)


# ═══════════════════════════════════════════════════════════════════════════
# Section 3 — Convert
# ═══════════════════════════════════════════════════════════════════════════

class ConvertSection(ttk.LabelFrame):
    def __init__(self, parent) -> None:
        super().__init__(parent, text="  Step 3 — Batch Convert  ", padding=10)
        self._targets_path = DEFAULT_TARGETS_JSON
        self._summary_var = tk.StringVar(value="No targets loaded.")
        self._status_var  = tk.StringVar(value="Idle")
        self._running = False
        self._abort_requested = False
        self._stop_event = threading.Event()
        self._log_q: queue.Queue = queue.Queue()
        self._log_handler: logging.Handler | None = None
        self._build_ui()
        if DEFAULT_TARGETS_JSON.exists():
            self._load_targets(DEFAULT_TARGETS_JSON)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # ── summary + status row ──────────────────────────────────────────
        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        top.columnconfigure(0, weight=1)

        info = ttk.Frame(top)
        info.grid(row=0, column=0, sticky="ew")
        info.columnconfigure(1, weight=1)

        ttk.Label(info, text="Targets:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(info, textvariable=self._summary_var).grid(
            row=0, column=1, sticky="w")

        ttk.Button(info, text="Browse…",
                   command=self._browse_targets, width=9).grid(
            row=0, column=2, sticky="e", padx=(8, 0))
        ttk.Button(info, text="Reload",
                   command=lambda: self._load_targets(
                       Path(self._targets_path)),
                   width=8).grid(row=0, column=3, padx=(6, 0))

        # ── control buttons ───────────────────────────────────────────────
        ctrl = ttk.Frame(self)
        ctrl.grid(row=1, column=0, sticky="ew", pady=(0, 6))

        self._run_btn = ttk.Button(ctrl, text="▶  Run Convert",
                                   command=self._start, width=16)
        self._run_btn.pack(side="left")

        self._stop_btn = ttk.Button(ctrl, text="■  Stop",
                                    command=self._stop, state="disabled", width=10)
        self._stop_btn.pack(side="left", padx=(8, 0))

        self._abort_btn = ttk.Button(ctrl, text="✕  Abort",
                                     command=self._abort, state="disabled", width=10)
        self._abort_btn.pack(side="left", padx=(6, 0))

        ttk.Separator(ctrl, orient="vertical").pack(
            side="left", fill="y", padx=14, pady=2)

        ttk.Button(ctrl, text="Clear log",
                   command=self._clear_log).pack(side="left")

        ttk.Separator(ctrl, orient="vertical").pack(
            side="left", fill="y", padx=14, pady=2)

        ttk.Button(ctrl, text="⚙ Config",
                   command=self._open_config, width=10).pack(side="left")

        ttk.Separator(ctrl, orient="vertical").pack(
            side="left", fill="y", padx=14, pady=2)

        ttk.Label(ctrl, text="Status:").pack(side="left", padx=(0, 6))
        self._status_lbl = ttk.Label(ctrl, textvariable=self._status_var,
                                     foreground="gray", width=20)
        self._status_lbl.pack(side="left")

        self._progress = ttk.Progressbar(ctrl, mode="indeterminate",
                                         length=140, maximum=100)
        self._progress.pack(side="left", padx=(10, 0))

        self._progress_lbl = ttk.Label(ctrl, text="", width=10,
                                       foreground="gray")
        self._progress_lbl.pack(side="left", padx=(4, 0))

        # ── log ───────────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self, text="Log", padding=4)
        log_frame.grid(row=2, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap="word", state="disabled",
            font=("Consolas", 9), height=12,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.tag_config("ERR",  foreground="#cc0000")
        self.log_text.tag_config("WARN", foreground="#cc6600")
        self.log_text.tag_config("OK",   foreground="#007700")

    # ── callbacks ──────────────────────────────────────────────────────────

    def load_targets(self, json_path: Path) -> None:
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
        if not json_path.exists():
            self._summary_var.set("File not found.")
            return
        try:
            targets = json.loads(json_path.read_text(encoding="utf-8"))
            optv = sum(1 for t in targets if t.get("data_type") == "OPTV")
            bhtv = sum(1 for t in targets if t.get("data_type") == "BHTV")
            self._summary_var.set(
                f"{len(targets)} targets loaded — {optv} OPTV, {bhtv} BHTV"
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
        elif " OK" in upper or " DONE" in upper:
            tag = "OK"
        self.log_text.insert("end", text + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _update_progress(self, current: int, total: int, label: str) -> None:
        if self._progress["mode"] == "indeterminate":
            self._progress.stop()
            self._progress.configure(mode="determinate")
        pct = int(current / total * 100) if total else 0
        self._progress.configure(maximum=100, value=pct)
        self._progress_lbl.configure(text=f"{current}/{total}")
        self._status_var.set(label.split("—")[-1].strip() if "—" in label else "Running…")

    def _poll_log(self) -> None:
        try:
            while True:
                msg = self._log_q.get_nowait()
                if msg is None:
                    self._on_done()
                    return
                if "__PROGRESS__" in msg:
                    try:
                        part = msg.split("__PROGRESS__")[1].strip()
                        n_str, rest = part.split("/", 1)
                        total_str = rest.split(" ")[0]
                        label = rest[len(total_str):].strip()
                        self._update_progress(int(n_str), int(total_str), label)
                    except Exception:
                        pass
                    continue   # don't write progress tokens to the log widget
                self._append_log(msg)
        except queue.Empty:
            pass
        if self._running:
            self.after(100, self._poll_log)

    def _set_running(self, running: bool) -> None:
        self._running = running
        if running:
            self._run_btn.configure(state="disabled")
            self._stop_btn.configure(state="normal")
            self._abort_btn.configure(state="normal")
            self._status_var.set("Running…")
            self._status_lbl.configure(foreground="#cc6600")
            self._progress.configure(mode="indeterminate")
            self._progress.start(12)
            self._progress_lbl.configure(text="")
        else:
            self._run_btn.configure(state="normal")
            self._stop_btn.configure(state="disabled")
            self._abort_btn.configure(state="disabled")
            self._progress.stop()

    def _start(self) -> None:
        if self._running:
            return
        if not self._targets_path.exists():
            messagebox.showerror("No targets",
                                 f"File not found:\n{self._targets_path}")
            return
        self._set_running(True)
        self._abort_requested = False
        self._stop_event.clear()
        self._append_log(f"=== Starting convert: {self._targets_path} ===")

        root_log = logging.getLogger()
        root_log.setLevel(logging.INFO)
        handler = _QueueHandler(self._log_q)
        handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)-8s %(message)s"))
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
        self._stop_btn.configure(state="disabled")
        self._status_var.set("Stopping…")
        self._status_lbl.configure(foreground="#cc6600")

    def _abort(self) -> None:
        self._stop_event.set()
        self._abort_requested = True
        self._append_log("=== ABORT — killing vendor app now ===")
        self._stop_btn.configure(state="disabled")
        self._abort_btn.configure(state="disabled")
        self._status_var.set("Aborting…")
        self._status_lbl.configure(foreground="#cc0000")

        def _kill() -> None:
            try:
                from modules.convert.batch_convert import _kill_app, OPTV_PATH, BHTV_PATH
                _kill_app(OPTV_PATH)
                _kill_app(BHTV_PATH)
                self._log_q.put("Vendor app killed.")
            except Exception as exc:
                self._log_q.put(f"Abort kill error: {exc}")

        threading.Thread(target=_kill, daemon=True).start()

    def _open_config(self) -> None:
        _DEFAULT_OPTV = r"C:\Electromind\OPTV Logger\OPTV.exe"
        _DEFAULT_BHTV = r"C:\Electromind\BHTV Logger\BHTV.exe"
        state = _load_state()
        optv_var = tk.StringVar(value=state.get("optv_path", _DEFAULT_OPTV))
        bhtv_var = tk.StringVar(value=state.get("bhtv_path", _DEFAULT_BHTV))

        dlg = tk.Toplevel(self)
        dlg.title("Config — Vendor Paths")
        dlg.resizable(False, False)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)

        def _browse_exe(var: tk.StringVar) -> None:
            current = var.get()
            init_dir = str(Path(current).parent) if current else str(Path.home())
            f = filedialog.askopenfilename(
                title="Select executable",
                filetypes=[("Executable", "*.exe"), ("All", "*.*")],
                initialdir=init_dir,
            )
            if f:
                var.set(f)

        for row, (label, var) in enumerate([
            ("OPTV.exe:", optv_var),
            ("BHTV.exe:", bhtv_var),
        ]):
            ttk.Label(frame, text=label, width=10, anchor="w").grid(
                row=row, column=0, sticky="w", padx=(0, 6), pady=4)
            ttk.Entry(frame, textvariable=var, width=60).grid(
                row=row, column=1, sticky="ew", pady=4)
            ttk.Button(frame, text="Browse…",
                       command=lambda v=var: _browse_exe(v), width=9).grid(
                row=row, column=2, padx=(6, 0), pady=4)

        btn_frame = ttk.Frame(dlg, padding=(12, 0, 12, 12))
        btn_frame.grid(row=1, column=0, sticky="ew")

        def _ok() -> None:
            _save_state("optv_path", optv_var.get().strip())
            _save_state("bhtv_path", bhtv_var.get().strip())
            dlg.destroy()

        ttk.Button(btn_frame, text="OK", command=_ok, width=10).pack(side="right")
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy,
                   width=10).pack(side="right", padx=(0, 6))

        dlg.wait_window()

    def _on_done(self) -> None:
        self._running = False
        self._set_running(False)
        if self._log_handler and hasattr(self, "_root_log"):
            self._root_log.removeHandler(self._log_handler)
            self._log_handler = None
        self._append_log("=== Convert complete ===")
        if self._abort_requested:
            self._status_var.set("Aborted")
            self._status_lbl.configure(foreground="#cc0000")
        elif self._stop_event.is_set():
            self._status_var.set("Stopped")
            self._status_lbl.configure(foreground="#cc6600")
        else:
            self._status_var.set("Done")
            self._status_lbl.configure(foreground="#007700")
        # Fill bar to 100% on completion
        self._progress.configure(mode="determinate", value=100)
        self._progress_lbl.configure(text="complete")


# ═══════════════════════════════════════════════════════════════════════════
# Main application
# ═══════════════════════════════════════════════════════════════════════════

class LauncherApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Televiewer Pipeline")
        self.geometry("1000x760")
        self.minsize(900, 720)

        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass

        pw = ttk.PanedWindow(self, orient="vertical")
        pw.pack(fill="both", expand=True, padx=10, pady=10)

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

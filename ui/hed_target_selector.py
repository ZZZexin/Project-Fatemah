"""Tkinter utility for selecting HED target files per hole.

The app scans a chosen directory recursively for .hed files, groups them by
hole name, and lets the user pick the desired OPTV and BHTV run for each hole.
Selections are saved as JSON/CSV so later pipeline steps can consume the chosen
target paths.
"""

from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from organize_hole_files import apply_plan, build_plan
except ModuleNotFoundError:
    from ui.organize_hole_files import apply_plan, build_plan


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SEARCH_DIR = Path(r"C:\2026_RTIO\West Angelas\Sent")
DEFAULT_JSON_OUTPUT = PROJECT_ROOT / "config" / "selected_hed_targets.json"
DEFAULT_CSV_OUTPUT = PROJECT_ROOT / "config" / "selected_hed_targets.csv"


@dataclass(frozen=True)
class HedRun:
    hole: str
    data_type: str
    label: str
    path: Path


OPTV_MARKERS = ("OPTV", "OTV", "OBI")
BHTV_MARKERS = ("BHTV", "ATV", "ABI", "BTV")
SELECTABLE_TYPES = ("OPTV", "BHTV")
EMPTY_CHOICE = ""
ALL_CHOICE = "ALL"


def has_marker(path: Path, markers: tuple[str, ...]) -> bool:
    text = " ".join((path.stem, *path.parts)).upper()
    return any(marker in text for marker in markers)


def classify_run(path: Path) -> str | None:
    """Classify HED files using both folder names and file names."""
    if "LGX" in (part.upper() for part in path.parts) or "LGX" in path.stem.upper():
        return None
    if has_marker(path, BHTV_MARKERS):
        return "BHTV"
    if has_marker(path, OPTV_MARKERS):
        return "OPTV"
    return None


def infer_hole_name(path: Path, _search_root: Path) -> str:
    """Infer the hole from the first filename segment before an underscore."""
    stem = path.stem
    if "_" in stem:
        return stem.split("_", 1)[0]
    return stem


def build_run_label(path: Path, hole: str, search_root: Path) -> str:
    """Use the full .hed filename as the label — unambiguous and clear."""
    return path.name


def default_run_label(data_type: str, runs: list[HedRun]) -> str:
    """Pick a sensible default run for each image type."""
    if not runs:
        return EMPTY_CHOICE

    preferred_text = "in" if data_type == "OPTV" else "out"
    for run in runs:
        searchable = f"{run.label} {run.path.name}".lower()
        if preferred_text in searchable:
            return run.label
    return runs[0].label


def find_hed_runs(search_dir: Path) -> dict[str, dict[str, list[HedRun]]]:
    runs_by_hole: dict[str, dict[str, list[HedRun]]] = {}
    seen: set[str] = set()  # "hole:type:filename" — dedup copies in organised subdirs
    for path in sorted(search_dir.rglob("*")):
        if not path.is_file() or path.suffix.lower() != ".hed":
            continue

        data_type = classify_run(path)
        if data_type is None:
            continue

        hole = infer_hole_name(path, search_dir)
        dedup_key = f"{hole}:{data_type}:{path.name}"
        if dedup_key in seen:
            continue
        seen.add(dedup_key)

        run = HedRun(
            hole=hole,
            data_type=data_type,
            label=build_run_label(path, hole, search_dir),
            path=path,
        )
        hole_runs = runs_by_hole.setdefault(hole, {data_type: [] for data_type in SELECTABLE_TYPES})
        hole_runs.setdefault(data_type, []).append(run)

    return dict(sorted(runs_by_hole.items(), key=lambda item: item[0].lower()))


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.inner.bind("<Configure>", self._update_scroll_region)
        self.canvas.bind("<Configure>", self._resize_inner)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _update_scroll_region(self, _event: tk.Event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _resize_inner(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _on_mousewheel(self, event: tk.Event) -> None:
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class HedTargetSelector(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("HED Target Selector")
        self.geometry("1040x720")
        self.minsize(860, 520)

        self.search_dir = tk.StringVar(value=str(DEFAULT_SEARCH_DIR))
        self.status = tk.StringVar(value="Choose a folder and scan for .hed files.")
        self.run_lookup: dict[str, dict[str, list[HedRun]]] = {}
        self.include_holes: dict[str, tk.BooleanVar] = {}
        self.selected_labels: dict[tuple[str, str], tk.StringVar] = {}
        self.path_labels: dict[tuple[str, str], tk.StringVar] = {}

        self._build_ui()
        if DEFAULT_SEARCH_DIR.exists():
            self.scan()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=12)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Search folder").grid(row=0, column=0, padx=(0, 8), sticky="w")
        path_entry = ttk.Entry(top, textvariable=self.search_dir)
        path_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(top, text="Browse", command=self.browse).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(top, text="Scan", command=self.scan).grid(row=0, column=3, padx=(8, 0))

        actions = ttk.Frame(top)
        actions.grid(row=1, column=0, columnspan=4, pady=(10, 0), sticky="ew")
        ttk.Button(actions, text="Organize to sorted", command=self.organize_to_sorted).pack(side="left")
        ttk.Button(actions, text="Select all", command=self.select_all).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Deselect all", command=self.deselect_all).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Save selections", command=self.save).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Copy paths", command=self.copy_paths).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Clear", command=self.clear_rows).pack(side="left", padx=(8, 0))
        ttk.Label(actions, textvariable=self.status).pack(side="left", padx=(18, 0))

        self.table = ScrollableFrame(self)
        self.table.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

    def browse(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.search_dir.get() or str(Path.cwd()))
        if selected:
            self.search_dir.set(selected)
            self.scan()

    def scan(self) -> None:
        root = Path(self.search_dir.get()).expanduser()
        if not root.exists() or not root.is_dir():
            messagebox.showerror("Folder not found", f"Cannot scan:\n{root}")
            return

        self.run_lookup = find_hed_runs(root)
        self.render_rows()
        run_count = sum(len(runs) for type_runs in self.run_lookup.values() for runs in type_runs.values())
        self.status.set(
            f"Found {run_count} source .hed files across {len(self.run_lookup)} holes. LGX files excluded."
        )

    def organize_to_sorted(self) -> None:
        source_dir = Path(self.search_dir.get()).expanduser()
        if not source_dir.exists() or not source_dir.is_dir():
            messagebox.showerror("Folder not found", f"Cannot organize:\n{source_dir}")
            return

        output_dir = source_dir if source_dir.name.lower() == "sorted" else source_dir / "sorted"
        actions = build_plan(source_dir.resolve(), output_dir.resolve())
        if not actions:
            output_dir.mkdir(parents=True, exist_ok=True)
            self.search_dir.set(str(output_dir))
            self.scan()
            self.status.set(f"No HED files needed sorting. Switched to {output_dir}.")
            return

        try:
            apply_plan(actions, mode="copy")
        except OSError as exc:
            messagebox.showerror("Organize failed", f"Could not organize files:\n{exc}")
            return

        self.search_dir.set(str(output_dir))
        self.scan()
        self.status.set(f"Copied {len(actions)} HED files into {output_dir} and switched to it.")

    def clear_rows(self) -> None:
        for child in self.table.inner.winfo_children():
            child.destroy()
        self.include_holes.clear()
        self.selected_labels.clear()
        self.path_labels.clear()
        self.status.set("Cleared the current list.")

    def render_rows(self) -> None:
        self.clear_rows()

        headings = ("Include", "Hole", "OPTV / OTV / OBI", "BHTV / ATV / ABI / BTV")
        for column, text in enumerate(headings):
            label = ttk.Label(self.table.inner, text=text, font=("", 10, "bold"))
            label.grid(row=0, column=column, sticky="ew", padx=6, pady=(0, 6))

        self.table.inner.columnconfigure(0, weight=0, minsize=70)
        self.table.inner.columnconfigure(1, weight=0, minsize=140)
        self.table.inner.columnconfigure(2, weight=1, minsize=260)
        self.table.inner.columnconfigure(3, weight=1, minsize=260)

        for row_index, (hole, type_runs) in enumerate(self.run_lookup.items(), start=1):
            include = tk.BooleanVar(value=True)
            self.include_holes[hole] = include
            ttk.Checkbutton(self.table.inner, variable=include).grid(
                row=row_index, column=0, sticky="w", padx=6, pady=4
            )
            ttk.Label(self.table.inner, text=hole).grid(
                row=row_index, column=1, sticky="w", padx=6, pady=4
            )

            self.render_selector(row_index, hole, "OPTV", type_runs.get("OPTV", []), 2)
            self.render_selector(row_index, hole, "BHTV", type_runs.get("BHTV", []), 3)

    def render_selector(
        self,
        row_index: int,
        hole: str,
        data_type: str,
        runs: list[HedRun],
        combo_column: int,
    ) -> None:
        choices = [*[run.label for run in runs], ALL_CHOICE] if runs else []
        selected = tk.StringVar(value=default_run_label(data_type, runs))
        selected_path = tk.StringVar(value="")
        key = (hole, data_type)
        self.selected_labels[key] = selected
        self.path_labels[key] = selected_path
        if runs:
            self.after_idle(lambda hole_name=hole, run_type=data_type: self.update_target_path(hole_name, run_type))

        combo = ttk.Combobox(
            self.table.inner,
            textvariable=selected,
            values=choices,
            state="readonly" if runs else "disabled",
            width=28,
        )
        combo.grid(row=row_index, column=combo_column, sticky="ew", padx=6, pady=4)
        combo.bind(
            "<<ComboboxSelected>>",
            lambda _event, hole_name=hole, run_type=data_type: self.update_target_path(
                hole_name, run_type
            ),
        )

    def update_target_path(self, hole: str, data_type: str) -> None:
        key = (hole, data_type)
        selected = self.selected_labels[key].get()
        if selected == ALL_CHOICE:
            self.path_labels[key].set(ALL_CHOICE)
            return
        for run in self.run_lookup.get(hole, {}).get(data_type, []):
            if run.label == selected:
                self.path_labels[key].set(str(run.path))
                return
        self.path_labels[key].set("")

    def selected_runs_for(self, hole: str, data_type: str) -> list[HedRun]:
        key = (hole, data_type)
        selected = self.selected_labels[key].get()
        runs = self.run_lookup.get(hole, {}).get(data_type, [])
        if selected == ALL_CHOICE:
            return runs
        return [run for run in runs if run.label == selected]

    def select_all(self) -> None:
        for include in self.include_holes.values():
            include.set(True)
        self.status.set(f"Selected all {len(self.include_holes)} holes.")

    def deselect_all(self) -> None:
        for include in self.include_holes.values():
            include.set(False)
        self.status.set("Deselected all holes.")

    def get_selected_targets(self) -> list[dict[str, str]]:
        targets: list[dict[str, str]] = []
        for hole in sorted(self.run_lookup):
            include = self.include_holes.get(hole)
            if include is None or not include.get():
                continue
            for data_type in SELECTABLE_TYPES:
                self.update_target_path(hole, data_type)
                for run in self.selected_runs_for(hole, data_type):
                    targets.append(
                        {
                            "hole": hole,
                            "data_type": data_type,
                            "run": run.label,
                            "parent_directory": str(run.path.parent),
                            "file_name": run.path.name,
                            "path": str(run.path),
                        }
                    )
        return targets

    def save(self) -> None:
        targets = self.get_selected_targets()
        if not targets:
            messagebox.showwarning("Nothing to save", "Scan first and select at least one run.")
            return

        DEFAULT_JSON_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
        DEFAULT_JSON_OUTPUT.write_text(json.dumps(targets, indent=2), encoding="utf-8")

        with DEFAULT_CSV_OUTPUT.open("w", newline="", encoding="utf-8") as csv_file:
            writer = csv.DictWriter(
                csv_file,
                fieldnames=["hole", "data_type", "run", "parent_directory", "file_name", "path"],
            )
            writer.writeheader()
            writer.writerows(targets)

        self.status.set(f"Saved {len(targets)} target paths to {DEFAULT_JSON_OUTPUT}.")
        messagebox.showinfo(
            "Selections saved",
            f"Saved JSON:\n{DEFAULT_JSON_OUTPUT.resolve()}\n\nSaved CSV:\n{DEFAULT_CSV_OUTPUT.resolve()}",
        )

    def copy_paths(self) -> None:
        paths = []
        for target in self.get_selected_targets():
            path = target.get("path")
            if path:
                paths.append(path)
        if not paths:
            messagebox.showwarning("Nothing to copy", "Scan first and select at least one run.")
            return

        self.clipboard_clear()
        self.clipboard_append("\n".join(paths))
        self.status.set(f"Copied {len(paths)} target paths to clipboard.")


def main() -> None:
    app = HedTargetSelector()
    app.mainloop()


if __name__ == "__main__":
    main()

"""
Batch converter — drives OPTV.exe and BHTV.exe to export LGX + LAS files.

Usage (from project root):
    python -m modules.convert.batch_convert
    python -m modules.convert.batch_convert config/selected_hed_targets.json
"""

import json
import logging
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Callable

import psutil
from pywinauto import Desktop
from pywinauto.application import Application

from modules.convert.vendor_app import (
    open_picture_dialog, export_picture_file, close_picture_dialog,
    open_optv_las_dialog, export_optv_las_file, close_optv_las_dialog,
    open_bhtv_las_dialog, export_bhtv_las_file, close_bhtv_las_dialog,
)

OPTV_PATH = r"C:\Electromind\OPTV Logger\OPTV.exe"
BHTV_PATH = r"C:\Electromind\BHTV Logger\BHTV.exe"

log = logging.getLogger(__name__)
_BM_CLICK = 0x00F5
_STARTUP_WARNING_TIMEOUT = 0.8

# ---------------------------------------------------------------------------
# App lifecycle helpers
# ---------------------------------------------------------------------------

def _kill_app(exe_path: str):
    exe_path = Path(exe_path)
    exe_name = exe_path.name.lower()
    procs = []
    for proc in psutil.process_iter(["pid", "name", "exe"]):
        try:
            same_name = (proc.info["name"] or "").lower() == exe_name
            same_path = proc.info["exe"] and Path(proc.info["exe"]).resolve() == exe_path.resolve()
            if same_name or same_path:
                procs.append(proc)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    for p in procs:
        try:
            p.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    if procs:
        psutil.wait_procs(procs, timeout=5)
        time.sleep(0.2)


def dismiss_startup_warning(app_title: str, process_id: int | None = None) -> bool:
    """Dismiss the startup warning dialog without moving the mouse."""
    attempts = []
    if process_id is not None:
        attempts.append({"title": app_title, "class_name": "#32770", "process": process_id})
    attempts.append({"title": app_title, "class_name": "#32770"})

    for criteria in attempts:
        try:
            dlg = Desktop(backend="win32").window(**criteria)
            dlg.wait("exists visible ready", timeout=_STARTUP_WARNING_TIMEOUT)
            ok = dlg.child_window(title="OK", class_name="Button")
            ok.wait("enabled", timeout=0.5)
            ok.send_message(_BM_CLICK)
            return True
        except Exception:
            pass
    return False


def _app_process_id(app) -> int | None:
    process = getattr(app, "process", None)
    if callable(process):
        try:
            return process()
        except Exception:
            return None
    return process


def _start_optv():
    start = time.perf_counter()
    _kill_app(OPTV_PATH)
    app = Application(backend="uia").start(OPTV_PATH)
    dismiss_startup_warning("OPTV", _app_process_id(app))
    main = app.window(title_re=".*OPTV Acquisition.*")
    main.wait("visible", timeout=3)
    log.info("OPTV startup ready in %.1fs", time.perf_counter() - start)
    return main


def _start_bhtv():
    start = time.perf_counter()
    _kill_app(BHTV_PATH)
    app = Application(backend="uia").start(BHTV_PATH)
    dismiss_startup_warning("BHTV", _app_process_id(app))
    main = app.window(title_re=".*BHTV Acquisition.*")
    main.wait("visible", timeout=3)
    log.info("BHTV startup ready in %.1fs", time.perf_counter() - start)
    return main


# ---------------------------------------------------------------------------
# Generic batch loop — dialog opened once, reused for every file
# ---------------------------------------------------------------------------

def _run_loop(entries, app_start_fn, open_dialog_fn, export_fn, close_dialog_fn,
              kill_path, export_label, stop_event=None,
              progress_cb: Callable[[int, int, str], None] | None = None,
              progress_offset: int = 0, progress_total: int = 0):
    """
    Open dialog once, iterate entries calling export_fn(dialog, hed_path),
    close dialog, kill app.
    On per-file failure: log, restart app + re-open dialog, continue next entry.
    stop_event (threading.Event): checked before each file.
    progress_cb(current, total, label): called after each completed file.
    Returns {hole: "ok" | "FAILED: ..." | "SKIPPED: ..."}.
    """
    results = {}
    main = app_start_fn()
    dialog_start = time.perf_counter()
    dialog = open_dialog_fn(main)
    log.info("%s dialog ready in %.1fs", export_label, time.perf_counter() - dialog_start)

    for i, entry in enumerate(entries):
        if stop_event and stop_event.is_set():
            log.info("Stop requested — skipping remaining %d entries.", len(entries) - i)
            for e in entries[i:]:
                results[e["hole"]] = "SKIPPED: stopped by user"
            break

        hole = entry["hole"]
        hed_path = entry["path"]
        try:
            log.info("[%s] %s — start", hole, export_label)
            result = export_fn(dialog, hed_path)
            if result == "skipped":
                results[hole] = "SKIPPED: output exists"
                log.info("[%s] %s — skipped: output exists", hole, export_label)
            else:
                results[hole] = "ok"
                log.info("[%s] %s — done", hole, export_label)
            if progress_cb:
                progress_cb(progress_offset + i + 1, progress_total,
                            f"{export_label} — {hole}")
        except Exception as exc:
            results[hole] = f"FAILED: {exc}"
            log.error("[%s] %s — failed: %s  ->  restarting app", hole, export_label, exc)
            _kill_app(kill_path)
            time.sleep(2)
            try:
                main = app_start_fn()
                dialog = open_dialog_fn(main)
            except Exception as restart_exc:
                log.error("App restart failed: %s  ->  aborting remaining entries", restart_exc)
                for e in entries[i + 1:]:
                    results[e["hole"]] = "SKIPPED: app restart failed"
                break

    try:
        close_dialog_fn(dialog)
    except Exception:
        pass
    _kill_app(kill_path)
    return results


# ---------------------------------------------------------------------------
# Public batch runners
# ---------------------------------------------------------------------------

def run_optv_batch(entries: list, stop_event=None,
                   progress_cb=None, progress_offset: int = 0,
                   progress_total: int = 0) -> dict:
    """Picture loop then LAS loop for all OPTV entries."""
    n = len(entries)
    log.info("=== OPTV Picture export: %d holes ===", n)
    picture = _run_loop(
        entries, _start_optv,
        open_picture_dialog, export_picture_file, close_picture_dialog,
        OPTV_PATH, "OPTV picture", stop_event,
        progress_cb=progress_cb,
        progress_offset=progress_offset,
        progress_total=progress_total,
    )

    if stop_event and stop_event.is_set():
        holes = {e["hole"] for e in entries}
        return {hole: {"picture": picture.get(hole),
                       "las": "SKIPPED: stopped by user"} for hole in holes}

    log.info("=== OPTV LAS export: %d holes ===", n)
    las = _run_loop(
        entries, _start_optv,
        open_optv_las_dialog, export_optv_las_file, close_optv_las_dialog,
        OPTV_PATH, "OPTV las", stop_event,
        progress_cb=progress_cb,
        progress_offset=progress_offset + n,
        progress_total=progress_total,
    )

    holes = {e["hole"] for e in entries}
    return {hole: {"picture": picture.get(hole), "las": las.get(hole)} for hole in holes}


def run_bhtv_batch(entries: list, stop_event=None,
                   progress_cb=None, progress_offset: int = 0,
                   progress_total: int = 0) -> dict:
    """LAS loop for all BHTV entries."""
    log.info("=== BHTV LAS export: %d holes ===", len(entries))
    las = _run_loop(
        entries, _start_bhtv,
        open_bhtv_las_dialog, export_bhtv_las_file, close_bhtv_las_dialog,
        BHTV_PATH, "BHTV las", stop_event,
        progress_cb=progress_cb,
        progress_offset=progress_offset,
        progress_total=progress_total,
    )
    return {hole: {"las": result} for hole, result in las.items()}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def run_all(targets_json_path: str = "config/selected_hed_targets.json",
            stop_event=None, progress_cb=None):
    global OPTV_PATH, BHTV_PATH

    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / f"convert_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    formatter = logging.Formatter("%(asctime)s %(levelname)-8s %(message)s")
    root_log = logging.getLogger()
    had_handlers = bool(root_log.handlers)
    root_log.setLevel(logging.INFO)
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(formatter)
    root_log.addHandler(file_handler)
    if not had_handlers:
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        root_log.addHandler(stream_handler)

    # Load configured exe paths from persisted state (set via Config dialog)
    if getattr(sys, "frozen", False):
        _state_path = Path(sys.executable).resolve().parent / "config" / "launcher_state.json"
    else:
        _state_path = Path(__file__).resolve().parents[2] / "config" / "launcher_state.json"
    try:
        _st = json.loads(_state_path.read_text(encoding="utf-8"))
        if _st.get("optv_path"):
            OPTV_PATH = _st["optv_path"]
        if _st.get("bhtv_path"):
            BHTV_PATH = _st["bhtv_path"]
    except Exception:
        pass

    with open(targets_json_path) as f:
        targets = json.load(f)

    optv = [t for t in targets if t["data_type"] == "OPTV"]
    bhtv = [t for t in targets if t["data_type"] == "BHTV"]
    log.info("Loaded %d OPTV, %d BHTV targets from %s", len(optv), len(bhtv), targets_json_path)

    try:
        # Total operations: OPTV needs picture + LAS (2 passes), BHTV needs LAS only
        total_ops = 2 * len(optv) + len(bhtv)
        _completed = [0]
        _lock = threading.Lock()

        def _progress(current: int, total: int, label: str):
            # Emit a special token the launcher UI detects for progress bar updates
            log.info("__PROGRESS__ %d/%d %s", current, total, label)

        # Build per-stream progress callbacks with thread-safe shared counter
        def _make_cb(offset: int):
            def cb(current: int, total: int, label: str):
                with _lock:
                    _completed[0] += 1
                    _progress(_completed[0], total_ops, label)
            return cb

        all_results = {}

        # Keep UI automation sequential so vendor apps do not compete for desktop focus.
        if optv:
            all_results["OPTV"] = run_optv_batch(
                optv, stop_event,
                progress_cb=_make_cb(0),
                progress_offset=0,
                progress_total=total_ops,
            )
        if bhtv and not (stop_event and stop_event.is_set()):
            bhtv_offset = 2 * len(optv)
            all_results["BHTV"] = run_bhtv_batch(
                bhtv, stop_event,
                progress_cb=_make_cb(bhtv_offset),
                progress_offset=bhtv_offset,
                progress_total=total_ops,
            )
    finally:
        if sys.exc_info()[0] is not None:
            root_log.removeHandler(file_handler)
            file_handler.close()

    log.info("=== SUMMARY ===")
    for dtype, results in all_results.items():
        for hole, status in results.items():
            for export_type, result in status.items():
                tag = "OK  " if result == "ok" else "SKIP" if str(result).startswith("SKIPPED") else "FAIL"
                log.info("  %s  [%s] %-20s %s  %s", tag, dtype, hole, export_type,
                         "" if result == "ok" else result)

    log.info("Log saved to %s", log_file)
    root_log.removeHandler(file_handler)
    file_handler.close()
    return all_results


if __name__ == "__main__":
    json_path = sys.argv[1] if len(sys.argv) > 1 else "config/selected_hed_targets.json"
    run_all(json_path)

"""
Batch converter — drives OPTV.exe and BHTV.exe to export LGX + LAS files.

Usage (from project root):
    python -m modules.convert.batch_convert
    python -m modules.convert.batch_convert config/selected_hed_targets.json
"""

import json
import logging
import sys
import time
from datetime import datetime
from pathlib import Path

import psutil
from pywinauto.application import Application

from modules.convert.vendor_app import export_optv_picture, export_optv_las, export_bhtv_las

OPTV_PATH = r"C:\Electromind\OPTV Logger\OPTV.exe"
BHTV_PATH = r"C:\Electromind\BHTV Logger\BHTV.exe"

log = logging.getLogger(__name__)


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
    time.sleep(1)


def _start_optv():
    _kill_app(OPTV_PATH)
    app = Application(backend="uia").start(OPTV_PATH)
    time.sleep(2)
    app.OPTV.OK.click()
    main = app.window(title_re=".*OPTV Acquisition.*")
    main.wait("visible", timeout=15)
    main.set_focus()
    return main


def _start_bhtv():
    _kill_app(BHTV_PATH)
    app = Application(backend="uia").start(BHTV_PATH)
    time.sleep(2)
    app.BHTV.OK.click()
    main = app.window(title_re=".*BHTV Acquisition.*")
    main.wait("visible", timeout=15)
    main.set_focus()
    return main


# ---------------------------------------------------------------------------
# Generic batch loop
# ---------------------------------------------------------------------------

def _run_loop(entries, export_fn, app_start_fn, kill_path, export_label):
    """
    Iterate entries, call export_fn(main, hed_path) for each.
    On failure: log, restart app, continue with next entry.
    Returns dict of {hole: "ok" | "FAILED: ..." | "SKIPPED: ..."}.
    """
    results = {}
    main = app_start_fn()

    for i, entry in enumerate(entries):
        hole = entry["hole"]
        hed_path = entry["path"]
        try:
            log.info("[%s] %s — start", hole, export_label)
            export_fn(main, hed_path)
            results[hole] = "ok"
            log.info("[%s] %s — done", hole, export_label)
        except Exception as exc:
            results[hole] = f"FAILED: {exc}"
            log.error("[%s] %s — failed: %s  →  restarting app", hole, export_label, exc)
            _kill_app(kill_path)
            time.sleep(2)
            try:
                main = app_start_fn()
            except Exception as restart_exc:
                log.error("App restart failed: %s  →  aborting remaining entries", restart_exc)
                for entry in entries[i + 1:]:
                    results[entry["hole"]] = "SKIPPED: app restart failed"
                break

    _kill_app(kill_path)
    return results


# ---------------------------------------------------------------------------
# Public batch runners
# ---------------------------------------------------------------------------

def run_optv_batch(entries: list) -> dict:
    """Picture loop then LAS loop for all OPTV entries."""
    log.info("=== OPTV Picture export: %d holes ===", len(entries))
    picture = _run_loop(entries, export_optv_picture, _start_optv, OPTV_PATH, "picture")

    log.info("=== OPTV LAS export: %d holes ===", len(entries))
    las = _run_loop(entries, export_optv_las, _start_optv, OPTV_PATH, "las")

    holes = {e["hole"] for e in entries}
    return {hole: {"picture": picture.get(hole), "las": las.get(hole)} for hole in holes}


def run_bhtv_batch(entries: list) -> dict:
    """LAS loop for all BHTV entries."""
    log.info("=== BHTV LAS export: %d holes ===", len(entries))
    las = _run_loop(entries, export_bhtv_las, _start_bhtv, BHTV_PATH, "las")
    return {hole: {"las": result} for hole, result in las.items()}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def run_all(targets_json_path: str = "config/selected_hed_targets.json"):
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / f"convert_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_file),
        ],
    )

    with open(targets_json_path) as f:
        targets = json.load(f)

    optv = [t for t in targets if t["data_type"] == "OPTV"]
    bhtv = [t for t in targets if t["data_type"] == "BHTV"]
    log.info("Loaded %d OPTV, %d BHTV targets from %s", len(optv), len(bhtv), targets_json_path)

    all_results = {}
    if optv:
        all_results["OPTV"] = run_optv_batch(optv)
    if bhtv:
        all_results["BHTV"] = run_bhtv_batch(bhtv)

    # Summary
    log.info("=== SUMMARY ===")
    for dtype, results in all_results.items():
        for hole, status in results.items():
            for export_type, result in status.items():
                tag = "OK  " if result == "ok" else "FAIL"
                log.info("  %s  [%s] %-20s %s  %s", tag, dtype, hole, export_type, "" if result == "ok" else result)

    log.info("Log saved to %s", log_file)
    return all_results


if __name__ == "__main__":
    # Allow running as: python -m modules.convert.batch_convert [path/to/targets.json]
    json_path = sys.argv[1] if len(sys.argv) > 1 else "config/selected_hed_targets.json"
    run_all(json_path)

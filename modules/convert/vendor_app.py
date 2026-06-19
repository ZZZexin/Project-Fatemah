"""Per-file export helpers for OPTV and BHTV vendor software.

Design: dialogs are opened ONCE before the batch loop and closed after.
Each export_*_file() function loads a new .hed into the already-open dialog.
click() is used instead of click_input() where possible so automation targets
controls directly without moving the mouse cursor.
"""

from __future__ import annotations

import time

import pyperclip
from pywinauto import Desktop

LAS_TARGETS = ["Inclination", "Azimuth", "Total Mag.", "Natural Gamma"]

_SWP_NOSIZE   = 0x0001
_SWP_NOZORDER = 0x0004


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _load_hed(hed_path: str):
    """Set hed_path in the open 'Open' file dialog and confirm."""
    open_dialog = Desktop(backend="win32").window(title="Open", class_name="#32770")
    open_dialog.wait("exists visible ready", timeout=10)

    try:
        edit = open_dialog.child_window(class_name="Edit")
        edit.wait("exists enabled", timeout=3)
        edit.set_edit_text(hed_path)
        open_dialog.child_window(title_re=".*Open.*", class_name="Button").click()
        return
    except Exception:
        pass

    # Fallback for older common dialogs that reject direct edit-control updates.
    open_dialog.set_focus()
    pyperclip.copy(hed_path)
    open_dialog.type_keys("%n")
    time.sleep(0.05)
    open_dialog.type_keys("^a")
    open_dialog.type_keys("^v")
    time.sleep(0.05)
    open_dialog.type_keys("{ENTER}")


def _ensure_checked(parent_window, title: str):
    cb = parent_window.child_window(title=title, class_name="Button")
    cb.wait("exists enabled", timeout=2)
    if cb.get_check_state() == 0:
        cb.click()


def _click_browse_button(dialog) -> None:
    """Click the dialog browse button, preferring named controls over index order."""
    for title_re in (".*Browse.*", r".*\.\.\..*"):
        try:
            btn = dialog.child_window(title_re=title_re, class_name="Button")
            btn.wait("exists enabled", timeout=1)
            btn.click()
            return
        except Exception:
            pass

    dialog.descendants()[3].click()


def _handle_dialog(timeout: float = 1.0) -> str:
    """
    Dismiss any OPTV or BHTV-titled popup dialog.

    Tries No first (replace/skip dialog) then OK (error or completion dialog).
    Returns: 'skipped' | 'ok' | 'none'
    """
    for title in ("OPTV", "BHTV"):
        try:
            dlg = Desktop(backend="win32").window(title=title, class_name="#32770")
            dlg.wait("exists visible ready", timeout=timeout)
        except Exception:
            continue

        try:
            no_btn = dlg.child_window(title_re=".*No.*", class_name="Button")
            no_btn.wait("enabled", timeout=1)
            no_btn.click()
            return "skipped"
        except Exception:
            pass

        try:
            ok_btn = dlg.child_window(title_re=".*OK.*", class_name="Button")
            ok_btn.wait("enabled", timeout=1)
            ok_btn.click()
            return "ok"
        except Exception:
            pass

    return "none"


def _poll_handle_dialog():
    """
    Non-blocking popup check used inside monitoring loops.
    Only handles the dialog if it already exists — no waiting.
    """
    for title in ("OPTV", "BHTV"):
        try:
            dlg = Desktop(backend="win32").window(title=title, class_name="#32770")
            if not dlg.exists():
                continue
        except Exception:
            continue

        try:
            no_btn = dlg.child_window(title_re=".*No.*", class_name="Button")
            if no_btn.exists():
                no_btn.click()
                return "skipped"
        except Exception:
            pass

        try:
            ok_btn = dlg.child_window(title_re=".*OK.*", class_name="Button")
            if ok_btn.exists():
                ok_btn.click()
                return "ok"
        except Exception:
            pass

    return "none"


def _has_lgx_pass(dialog) -> bool:
    """True if any direct child shows 'Creating LGX Pass' AND 'wait' — actively exporting."""
    try:
        for ctrl in dialog.children():
            try:
                t = ctrl.window_text()
                if "Creating LGX Pass" in t and "wait" in t.lower():
                    return True
            except Exception:
                pass
    except Exception:
        pass
    return False


def _wait_picture_complete(dialog, timeout: float = 300):
    """
    Wait for picture/LGX export to finish using a two-state machine.

    status 0→1 when 'Creating LGX Pass...wait...' appears in a direct child.
    status 1→0 (done) when it disappears and is still gone after 0.1 s.
    If it never appears within 15 s, assumes export was instant and returns.
    _poll_handle_dialog is called every 5th iteration (~0.5 s) to catch errors.
    """
    deadline = time.time() + timeout
    start_deadline = time.time() + 15
    status = 0
    poll_counter = 0

    while time.time() < deadline:
        poll_counter += 1
        if poll_counter % 5 == 0:
            _poll_handle_dialog()

        found = _has_lgx_pass(dialog)

        if found:
            status = 1
        elif status == 1:
            time.sleep(0.1)
            if not _has_lgx_pass(dialog):
                return   # export complete
        elif status == 0 and time.time() > start_deadline:
            raise TimeoutError("Picture export did not start before timeout")

        time.sleep(0.1)

    raise TimeoutError("Picture export did not finish before timeout")


def _wait_las_complete(dialog, timeout: float = 120):
    """Wait for LAS export to finish. Returns as soon as 'File migred' appears in a direct child."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            for ctrl in dialog.children():
                try:
                    if "File migred" in ctrl.window_text():
                        return
                except Exception:
                    pass
        except Exception:
            pass
        time.sleep(0.1)
    raise TimeoutError("LAS export did not finish before timeout")


def _open_export_menu(main, shortcut_key: str):
    """Alt+F → x (Export) → shortcut_key."""
    menu_paths = {
        "p": ("File->Export->Picture", "File->Export->Picture export"),
        "l": ("File->Export->LAS", "File->Export->LAS 2.0"),
    }
    for path in menu_paths.get(shortcut_key, ()):
        try:
            main.menu_select(path)
            time.sleep(0.05)
            return
        except Exception:
            pass

    main.set_focus()
    main.type_keys("%f")
    time.sleep(0.05)
    main.type_keys("x")
    time.sleep(0.05)
    main.type_keys(shortcut_key)
    time.sleep(0.05)


# ---------------------------------------------------------------------------
# Picture export  (OPTV only)
# ---------------------------------------------------------------------------

def open_picture_dialog(main):
    """Navigate menu and return the open Picture export dialog."""
    _open_export_menu(main, "p")
    dialog = Desktop(backend="win32").window(title_re=".*Picture export.*")
    dialog.wait("exists visible ready", timeout=10)
    return dialog


def export_picture_file(dialog, hed_path: str):
    """Load one .hed and export picture/LGX. Dialog must already be open."""
    _click_browse_button(dialog)
    _load_hed(hed_path)


    for btn_title in ("True color", "Bicubic"):
        btn = dialog.child_window(title=btn_title, class_name="Button")
        if btn.get_check_state() == 0:
            btn.click()

    lgx_cb = dialog.child_window(title="LGX Format", class_name="Button")
    lgx_cb.wait("enabled", timeout=10)
    if lgx_cb.get_check_state() == 0:
        lgx_cb.click()
    time.sleep(0.1)

    dialog.child_window(title="&Export", class_name="Button").click()

    _handle_dialog(timeout=1.0)

    _wait_picture_complete(dialog, timeout=300)


def close_picture_dialog(dialog):
    dialog.child_window(title="&Cancel", class_name="Button").click()


# ---------------------------------------------------------------------------
# LAS export — OPTV
# ---------------------------------------------------------------------------

def open_optv_las_dialog(main):
    """Navigate menu and return the open LAS export dialog for OPTV."""
    _open_export_menu(main, "l")
    dialog = Desktop(backend="win32").window(
        title="LAS 2.0 exportation for OPTV", class_name="#32770"
    )
    dialog.wait("exists visible ready", timeout=10)
    return dialog


def export_optv_las_file(dialog, hed_path: str):
    """Load one .hed and export LAS. Dialog must already be open."""
    _click_browse_button(dialog)
    _load_hed(hed_path)

    for t in LAS_TARGETS:
        _ensure_checked(dialog, t)

    dialog.child_window(title="&Start", class_name="Button").click()

    result = _handle_dialog(timeout=1.0)
    if result == "skipped":
        _handle_dialog(timeout=2.0)
    else:
        _wait_las_complete(dialog, timeout=120)


def close_optv_las_dialog(dialog):
    dialog.child_window(title="&Close", class_name="Button").click()


# ---------------------------------------------------------------------------
# LAS export — BHTV
# ---------------------------------------------------------------------------

def open_bhtv_las_dialog(main):
    """Navigate menu and return the open LAS export dialog for BHTV."""
    _open_export_menu(main, "l")
    dialog = Desktop(backend="win32").window(
        title="LAS 2.0 exportation for BHTV", class_name="#32770"
    )
    dialog.wait("exists visible ready", timeout=10)
    return dialog


def export_bhtv_las_file(dialog, hed_path: str):
    """Load one .hed and export LAS. Dialog must already be open."""
    _click_browse_button(dialog)
    _load_hed(hed_path)

    for t in LAS_TARGETS:
        _ensure_checked(dialog, t)

    dialog.child_window(title="&Start", class_name="Button").click()

    result = _handle_dialog(timeout=1.0)
    if result == "skipped":
        _handle_dialog(timeout=2.0)
    else:
        _wait_las_complete(dialog, timeout=120)


def close_bhtv_las_dialog(dialog):
    dialog.child_window(title="&Close", class_name="Button").click()

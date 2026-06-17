from pywinauto import Desktop
import pyperclip
import time

LAS_TARGETS = ["Inclination", "Azimuth", "Total Mag.", "Natural Gamma"]


def _load_hed(hed_path: str):
    open_dialog = Desktop(backend="win32").window(title="Open", class_name="#32770")
    open_dialog.wait("exists visible ready", timeout=10)
    open_dialog.set_focus()
    pyperclip.copy(hed_path)
    open_dialog.type_keys("%n")
    time.sleep(0.2)
    open_dialog.type_keys("^a")
    open_dialog.type_keys("^v")
    time.sleep(0.2)
    open_dialog.type_keys("{ENTER}")


def _ensure_checked(parent_window, title: str):
    cb = parent_window.child_window(title=title, class_name="Button")
    cb.wait("exists enabled", timeout=10)
    if cb.get_check_state() == 0:
        cb.click_input()


def _try_replace_dialog(app_title: str, timeout: float = 3.0) -> bool:
    """Click No on the 'file already exists — replace?' dialog if it appears."""
    try:
        dlg = Desktop(backend="win32").window(title=app_title, class_name="#32770")
        dlg.wait("exists visible ready", timeout=timeout)
        dlg.set_focus()
        no_btn = dlg.child_window(title_re=".*No.*", class_name="Button")
        no_btn.wait("enabled", timeout=5)
        no_btn.click_input()
        return True
    except Exception:
        return False


def _click_ok_dialog(app_title: str, timeout: float = 60.0):
    dlg = Desktop(backend="win32").window(title=app_title, class_name="#32770")
    dlg.wait("exists visible ready", timeout=timeout)
    dlg.set_focus()
    ok_btn = dlg.child_window(title_re=".*OK.*", class_name="Button")
    ok_btn.wait("enabled", timeout=10)
    ok_btn.click_input()


def export_optv_picture(main, hed_path: str):
    """Picture/LGX export for one OPTV .hed file. App must already be open."""
    main.set_focus()
    main.type_keys("%f")
    time.sleep(0.1)
    main.type_keys("x")
    time.sleep(0.1)
    main.type_keys("p")
    time.sleep(0.1)

    dialog = Desktop(backend="win32").window(title_re=".*Picture export.*")
    dialog.wait("visible", timeout=10)
    dialog.set_focus()
    dialog.descendants()[0].click_input()
    _load_hed(hed_path)

    dialog = Desktop(backend="win32").window(title="Picture export", class_name="#32770")
    dialog.wait("exists visible ready", timeout=10)
    dialog.set_focus()

    for btn_title in ("True color", "Bicubic"):
        btn = dialog.child_window(title=btn_title, class_name="Button")
        if btn.get_check_state() == 0:
            btn.click_input()

    lgx_cb = dialog.child_window(title="LGX Format", class_name="Button")
    lgx_cb.wait("enabled", timeout=10)
    if lgx_cb.get_check_state() == 0:
        lgx_cb.click_input()
    time.sleep(0.3)

    export_btn = dialog.child_window(title="&Export", class_name="Button")
    export_btn.click_input()

    _try_replace_dialog("OPTV", timeout=3.0)
    # Button re-enables when dialog is ready again (either after skip or after actual export)
    export_btn.wait("enabled", timeout=120)

    dialog.child_window(title="&Cancel", class_name="Button").click_input()


def export_optv_las(main, hed_path: str):
    """LAS export for one OPTV .hed file. App must already be open."""
    main.set_focus()
    main.type_keys("%f")
    time.sleep(0.1)
    main.type_keys("x")
    time.sleep(0.1)
    main.type_keys("l")
    time.sleep(0.1)

    las_dialog = Desktop(backend="win32").window(
        title="LAS 2.0 exportation for OPTV", class_name="#32770"
    )
    las_dialog.descendants()[3].click_input()
    _load_hed(hed_path)

    for t in LAS_TARGETS:
        _ensure_checked(las_dialog, t)

    las_dialog.child_window(title="&Start", class_name="Button").click_input()

    _try_replace_dialog("OPTV", timeout=3.0)
    _click_ok_dialog("OPTV", timeout=60)

    las_dialog.child_window(title="&Close", class_name="Button").click_input()


def export_bhtv_las(main, hed_path: str):
    """LAS export for one BHTV .hed file. App must already be open."""
    main.set_focus()
    main.type_keys("%f")
    time.sleep(0.1)
    main.type_keys("x")
    time.sleep(0.1)
    main.type_keys("l")
    time.sleep(0.1)

    las_dialog = Desktop(backend="win32").window(
        title="LAS 2.0 exportation for BHTV", class_name="#32770"
    )
    las_dialog.descendants()[3].click_input()
    _load_hed(hed_path)

    for t in LAS_TARGETS:
        _ensure_checked(las_dialog, t)

    las_dialog.child_window(title="&Start", class_name="Button").click_input()

    # Note: replace dialog in BHTV app has title "OPTV" (vendor quirk from initial_code)
    _try_replace_dialog("OPTV", timeout=3.0)
    _click_ok_dialog("BHTV", timeout=60)

    las_dialog.child_window(title="&Close", class_name="Button").click_input()

from pywinauto.application import Application
from pywinauto import Desktop
import pyperclip
import time
import psutil
from pathlib import Path


def kill_existing_app(exe_path: str, timeout: int = 5):
    exe_path = Path(exe_path)
    exe_name = exe_path.name.lower()

    found_processes = []

    for proc in psutil.process_iter(["pid", "name", "exe"]):
        try:
            proc_name = (proc.info["name"] or "").lower()
            proc_exe = proc.info["exe"]

            same_name = proc_name == exe_name
            same_path = proc_exe and Path(proc_exe).resolve() == exe_path.resolve()

            if same_name or same_path:
                found_processes.append(proc)

        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    if not found_processes:
        print("No existing OPTV process found.")
        return

    print(f"Found {len(found_processes)} existing OPTV process(es). Closing...")

    # Try graceful terminate first
    for proc in found_processes:
        try:
            print(f"Terminating PID {proc.pid}")
            proc.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
            print(f"Could not terminate PID {proc.pid}: {e}")

    gone, alive = psutil.wait_procs(found_processes, timeout=timeout)

    # Force kill if still alive
    for proc in alive:
        try:
            print(f"Force killing PID {proc.pid}")
            proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
            print(f"Could not kill PID {proc.pid}: {e}")

    time.sleep(1)
    print("Existing OPTV process cleanup finished.")

def ensure_checked(parent_window, title):
    checkbox = parent_window.child_window(
        title=title,
        class_name="Button"
    )

    checkbox.wait("exists enabled", timeout=10)

    state = checkbox.get_check_state()
    print(f"{title}: state before = {state}")

    # 0 = unchecked, 1 = checked, 2 = indeterminate
    if state == 0:
        checkbox.click_input()
        print(f"{title}: checked")

    else:
        print(f"{title}: already checked, no change")

    return checkbox.get_check_state()


def _open_app(tool_path):
    ''' tool path is either OPTV or BHTV'''
    app = Application(backend="uia").start(tool_path)
    
    if "OPTV" in tool_path:
        app.OPTV.OK.click()
        alias = "OPTV"
    elif "BHTV" in tool_path:
        app.BHTV.OK.click()
        alias = "BHTV"
    else: KeyError("Which tool? I don't understand")
    title_re = f".*{alias} Acquisition.*"
    main = app.window(title_re=title_re)
    main.set_focus()

    return main, alias

def _go_to_task(main, alias, task_type, idle_time = 0.1):
    if task_type == "las":
        shortcut_key = "l"
        child_kwd = f"LAS 2.0 exportation for {alias}"
    elif task_type == "picture":
        shortcut_key = "p"
        child_kwd = "Picture export"
    class_name = "#32770"

    main.type_keys("%f")       # Alt + F, open File menu
    time.sleep(idle_time)

    main.type_keys("x") # move to Export
    time.sleep(idle_time)

    main.type_keys(shortcut_key)  # open Export submenu
    time.sleep(idle_time)

    dialog = Desktop(backend="win32").window(
    title=child_kwd,
    class_name=class_name
    )
    dialog.wait("exists visible ready", timeout=idle_time*5)
    dialog.set_focus()

    return dialog

def _if_exist():
    for w in Desktop(backend="win32").windows():
        try:
            if w.window_text() == "OPTV"|"BHTV" and w.class_name() == "#32770":
                return True
        except Exception:
            pass

    return False

def _convert_pic(dialog, hed_path, idle_time=10):
    # since pic only needed for OPTV so alias fixed to OPTV
    dialog.descendants()[3].click_input()
    open_dialog = Desktop(backend="win32").window(
    title="Open",
    class_name="#32770"
    )
    open_dialog.wait("exists visible ready", timeout=10)
    open_dialog.set_focus()
    pyperclip.copy(hed_path)

    # Focus File name field
    open_dialog.type_keys("%n")
    time.sleep(0.2)

    # Replace current *.hed with actual file path
    open_dialog.type_keys("^a")
    open_dialog.type_keys("^v")
    time.sleep(0.2)

    # Press Open
    open_dialog.type_keys("{ENTER}")

    dialog.set_focus()

    # tick options
    set_true_color = dialog.child_window(title="True color", class_name="Button")
    tc_state = set_true_color.get_check_state()
    if tc_state == 0:
        set_true_color.click_input(title="True color", class_name="Button")

    set_interp = dialog.child_window(title="Bicubic", class_name="Button")
    interp_state = set_interp.get_check_state()
    if interp_state == 0:
        set_interp.click_input(title="Bicubic", class_name="Button")
        
    lgx_checkbox = dialog.child_window(
        title="LGX Format",
        class_name="Button"
    )

    lgx_checkbox.wait("enabled", timeout=1)

    # Check current state: 0 = unchecked, 1 = checked
    state = lgx_checkbox.get_check_state()
    print("LGX state before:", state)

    if state == 0:
        lgx_checkbox.click_input()

    time.sleep(0.3)

    print("LGX state after:", lgx_checkbox.get_check_state())

    lgx_export = dialog.child_window(

        title = '&Export',
        class_name = "Button"
    )

    lgx_export.click_input()

    # already converted?
    check_existing = _if_exist()
    if check_existing:
        replace_dialog = Desktop(backend="win32").window(
            title="OPTV",
            class_name="#32770"
        )
        replace_dialog.wait("exists visible ready", timeout=10)
        replace_dialog.set_focus()

        no_btn = replace_dialog.child_window(
            title_re=".*No.*",
            class_name="Button"
        )

        no_btn.wait("enabled", timeout=10)
        no_btn.click_input()


    else:
        time.sleep(idle_time) 
        # if not exist just wait till it is done, need to find better solution rather than seting the definite time it will be too long or too short to complete
    
    return dialog # leave this end for loop

#exit_dialog = dialog.child_window(title="&Cancel", class_name="Button")
#exit_dialog.click_input() # leave if only one hole

def _convert_las(dialog, alias, hed_path, idle_time=4):
    dialog.descendants()[3].click_input()
    open_dialog = Desktop(backend="win32").window(
    title="Open",
    class_name="#32770"
    )
    open_dialog.wait("exists visible ready", timeout=10)
    open_dialog.set_focus()

    pyperclip.copy(hed_path)
    # Focus File name field
    open_dialog.type_keys("%n")
    time.sleep(0.2)

    # Replace current *.hed with actual file path
    open_dialog.type_keys("^a")
    open_dialog.type_keys("^v")
    time.sleep(0.2)

    
    # Press Open
    open_dialog.type_keys("{ENTER}")

    targets = [
        "Inclination",
        "Azimuth",
        "Total Mag.",
        "Natural Gamma",
    ]

    for title in targets:
        ensure_checked(dialog, title)

    las_export = dialog.child_window(

    title = '&Start',
    class_name = "Button"
    )
    las_export.click_input()

    # already converted?
    check_existing = _if_exist()
    if check_existing:
        replace_dialog = Desktop(backend="win32").window(
            title=alias,
            class_name="#32770"
        )
        replace_dialog.wait("exists visible ready", timeout=10)
        replace_dialog.set_focus()

        no_btn = replace_dialog.child_window(
            title_re=".*No.*",
            class_name="Button"
        )

        no_btn.wait("enabled", timeout=10)
        no_btn.click_input()
        replace_dialog2 = Desktop(backend="win32").window(
            title=alias,
            class_name="#32770"
        )

        replace_dialog2.set_focus()

        ok_btn = replace_dialog2.child_window(
            title_re=".*OK.*",
            class_name="Button"
        )

        ok_btn.wait("enabled", timeout=10)
        ok_btn.click_input()

    else:
        time.sleep(idle_time) 
        # if not exist just wait till it is done, need to find better solution rather than seting the definite time it will be too long or too short to complete
    
    return dialog # leave this end for loop


def batch_running(dialog, hed_pathlist, alias, task_type,
                  time_p = 10,
                  time_l = 4):
    '''
    alias = OPTV | BHTV
    task_type = las | picture
    '''
    
    for hed_path in hed_pathlist:
        if task_type == "picture":
            dialog =  _convert_pic(dialog, hed_path, idle_time=time_p)
        elif task_type == "las":
            dialog = _convert_las(dialog, alias, hed_path, idle_time=time_l)

    return dialog

def exit_dialog(dialog, task_type):
    title = "&Cancel" if task_type == "picture" else "&Close"
    dialog.child_window(title=title, class_name="Button")
    exit_dialog.click_input()

def optv_flow(hed_pathlist, PATH):
    kill_existing_app(PATH)
    main, alias = _open_app(PATH)

    tasks = ["las", "picture"]
    for task in tasks:
        dialog = _go_to_task(main, alias, task, idle_time = 0.1) # this should be as fast as it can since just clicking the tab
        dialog = batch_running(dialog, hed_pathlist, alias, task,
                                time_p = 10,
                                time_l = 4)
        exit_dialog(dialog, task)
    kill_existing_app(PATH)


def bhtv_flow(hed_pathlist, PATH):
    kill_existing_app(PATH)
    main, alias = _open_app(PATH)

    tasks = ["las"]
    for task in tasks:
        dialog = _go_to_task(main, alias, task, idle_time = 0.1) # this should be as fast as it can since just clicking the tab
        dialog = batch_running(dialog, hed_pathlist, alias, task,
                                time_p = 10,
                                time_l = 4)
        exit_dialog(dialog, task)
    kill_existing_app(PATH)
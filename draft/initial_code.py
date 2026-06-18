from pywinauto.application import Application
from pywinauto import Desktop
import pyperclip
import time
import psutil
from pathlib import Path

OPTV_PATH = r"C:\Electromind\OPTV Logger\OPTV.exe"
BHTV_PATH = r"C:\Electromind\BHTV Logger\BHTV.exe"

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

# open app
# kill existing OPTV first
kill_existing_app(OPTV_PATH)

app = Application(backend="uia").start(OPTV_PATH)


# a window will bring up telling you there is no dongle, click ok
app.OPTV.OK.click()

# switch to the window
main = app.window(title_re=".*OPTV Acquisition.*")
main.set_focus()


main.type_keys("%f")       # Alt + F, open File menu
time.sleep(0.1)

main.type_keys("x") # move to Export
time.sleep(0.1)

main.type_keys("p")  # open Export submenu
time.sleep(0.1)

dialog = Desktop(backend="win32").window(title_re=".*Picture export.*")
dialog.wait("visible", timeout=10)
dialog.set_focus()

dialog.descendants()[0].click_input() # click input file

hed_file = r"C:\2026_RTIO\West Angelas\Sent\GR26WAH0001\OTV\GR26WAH0001_OBI_INRUN.hed"

open_dialog = Desktop(backend="win32").window(
    title="Open",
    class_name="#32770"
)

open_dialog.wait("exists visible ready", timeout=10)
open_dialog.set_focus()

pyperclip.copy(hed_file)

# Focus File name field
open_dialog.type_keys("%n")
time.sleep(0.2)

# Replace current *.hed with actual file path
open_dialog.type_keys("^a")
open_dialog.type_keys("^v")
time.sleep(0.2)

# Press Open
open_dialog.type_keys("{ENTER}")

dialog = Desktop(backend="win32").window(
    title="Picture export",
    class_name="#32770"
)

dialog.wait("exists visible ready", timeout=10)
dialog.set_focus()

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

lgx_checkbox.wait("enabled", timeout=10)

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

exit_dialog = dialog.child_window(title="&Cancel", class_name="Button")
exit_dialog.click_input() # leave if only one hole

# now we go las
main.set_focus()


main.type_keys("%f")       # Alt + F, open File menu
time.sleep(0.1)

main.type_keys("x") # move to Export
time.sleep(0.1)

main.type_keys("l")  # open Export submenu
time.sleep(0.1)


las_dialog = Desktop(backend="win32").window(
    title="LAS 2.0 exportation for OPTV",
    class_name="#32770"
)

las_dialog.descendants()[3].click_input()

open_dialog = Desktop(backend="win32").window(
    title="Open",
    class_name="#32770"
)

open_dialog.wait("exists visible ready", timeout=10)
open_dialog.set_focus()

pyperclip.copy(hed_file)
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
    ensure_checked(las_dialog, title)


las_export = las_dialog.child_window(

    title = '&Start',
    class_name = "Button"
)


las_export.click_input()

# if there is existing file, skipped. if not go on generate.
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
replace_dialog2 = Desktop(backend="win32").window(
    title="OPTV",
    class_name="#32770"
)

replace_dialog2.set_focus()

ok_btn = replace_dialog2.child_window(
    title_re=".*OK.*",
    class_name="Button"
)

ok_btn.wait("enabled", timeout=10)
ok_btn.click_input()

exit_dialog = las_dialog.child_window(title="&Close", class_name="Button")
exit_dialog.click_input() # leave if only one hole

kill_existing_app(OPTV_PATH)



############### BHTV, only need to convert las
hed_file = r"C:\2026_RTIO\Hope Downs\Sent\GR26HD40003\ATV\GR26HD40003_OUTRUN_H.hed"

kill_existing_app(BHTV_PATH)

app = Application(backend="uia").start(BHTV_PATH)

app.BHTV.OK.click()

# switch to the window
main = app.window(title_re=".*BHTV Acquisition.*")
main.set_focus()


main.type_keys("%f")       # Alt + F, open File menu
time.sleep(0.1)

main.type_keys("x") # move to Export
time.sleep(0.1)

main.type_keys("l")  # open Export submenu
time.sleep(0.1)


las_dialog = Desktop(backend="win32").window(
    title="LAS 2.0 exportation for BHTV",
    class_name="#32770"
)

las_dialog.descendants()[3].click_input()

open_dialog = Desktop(backend="win32").window(
    title="Open",
    class_name="#32770"
)

open_dialog.wait("exists visible ready", timeout=10)
open_dialog.set_focus()

pyperclip.copy(hed_file)
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
    ensure_checked(las_dialog, title)


las_export = las_dialog.child_window(

    title = '&Start',
    class_name = "Button"
)


las_export.click_input()

# if there is existing file, skipped. if not go on generate.
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
replace_dialog2 = Desktop(backend="win32").window(
    title="BHTV",
    class_name="#32770"
)

replace_dialog2.set_focus()

ok_btn = replace_dialog2.child_window(
    title_re=".*OK.*",
    class_name="Button"
)

ok_btn.wait("enabled", timeout=10)
ok_btn.click_input()

exit_dialog = las_dialog.child_window(title="&Close", class_name="Button")
exit_dialog.click_input() # leave if only one hole

kill_existing_app(BHTV_PATH)
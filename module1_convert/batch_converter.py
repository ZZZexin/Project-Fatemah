from pywinauto.application import Application
from pywinauto import Desktop
import time

OPTV_PATH = r"C:\Electromind\OPTV Logger\OPTV.exe"
BHTV_PATH = r"C:\Electromind\BHTV Logger\BHTV.exe"

app = Application(backend="uia").start(OPTV_PATH)
time.sleep(3)

# Find the OPTV popup/window
opvt_window = Desktop(backend="uia").window(title="OPTV")
opvt_window.wait("visible", timeout=30)

print("Window found:", opvt_window.window_text())

# Click OK without calling print_control_identifiers()
ok_button = opvt_window.child_window(
    title="OK",
    auto_id="2",
    control_type="Button"
)

ok_button.wait("enabled", timeout=10)
ok_button.click_input()

print("OK clicked.")
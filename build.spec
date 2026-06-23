# build.spec — PyInstaller spec for TvPipeline.exe
# Run:  pyinstaller build.spec --noconfirm --clean

from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# Collect everything from packages that use dynamic/late imports
datas_pw,  binaries_pw,  hiddenimports_pw  = collect_all("pywinauto")
datas_ct,  binaries_ct,  hiddenimports_ct  = collect_all("comtypes")
datas_psc, binaries_psc, hiddenimports_psc = collect_all("psutil")

a = Analysis(
    ["ui/launcher.py"],
    pathex=["."],
    binaries=binaries_pw + binaries_ct + binaries_psc,
    datas=datas_pw + datas_ct + datas_psc,
    hiddenimports=(
        hiddenimports_pw
        + hiddenimports_ct
        + hiddenimports_psc
        + [
            "pyperclip",
            "pyperclip.clipboards",
            "modules.convert.batch_convert",
            "modules.convert.vendor_app",
            "ui.organize_hole_files",
            "ui.hed_target_selector",
            "organize_hole_files",
            "hed_target_selector",
            "tkinter",
            "tkinter.ttk",
            "tkinter.filedialog",
            "tkinter.messagebox",
            "tkinter.scrolledtext",
        ]
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Trim heavy packages that are definitely not used
    excludes=["matplotlib", "numpy", "pandas", "scipy", "PIL", "cv2",
              "notebook", "IPython", "pytest"],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="TvPipeline",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX can trigger AV false positives — leave off
    runtime_tmpdir=None,
    console=False,      # no console window (pure GUI app)
    disable_windowed_traceback=False,
    icon=None,
)

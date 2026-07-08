# PyInstaller spec for the one-click WSR Generator desktop app.
# Build:  pyinstaller --noconfirm WSR-Generator.spec  ->  dist/WSR-Generator.exe
from PyInstaller.utils.hooks import collect_data_files

# Read-only assets that must travel inside the executable.
datas = [
    ("templates/CES_CSAR_WSR_Template.pptx", "templates"),
    ("assets/closing_backdrop.png", "assets"),
]
datas += collect_data_files("pptx")

a = Analysis(
    ["wsr_app.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="WSR-Generator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

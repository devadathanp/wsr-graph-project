# PyInstaller spec for the one-click WSR Generator desktop app.
# Windows:  pyinstaller --noconfirm WSR-Generator.spec  ->  dist/WSR-Generator.exe (single file)
# macOS:    pyinstaller --noconfirm WSR-Generator.spec  ->  dist/WSR-Generator.app (bundle)
import sys

from PyInstaller.utils.hooks import collect_data_files

IS_MAC = sys.platform == "darwin"

# Read-only assets that must travel inside the app.
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

if IS_MAC:
    # One-dir build wrapped in a .app bundle (recommended for macOS security).
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name="WSR-Generator",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name="WSR-Generator",
    )
    app = BUNDLE(
        coll,
        name="WSR-Generator.app",
        icon=None,
        bundle_identifier="com.kpit.wsrgenerator",
        info_plist={
            "CFBundleName": "WSR Generator",
            "CFBundleDisplayName": "WSR Generator",
            "CFBundleShortVersionString": "1.0.0",
            "NSHighResolutionCapable": True,
        },
    )
else:
    # Single-file executable for Windows.
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

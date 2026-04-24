# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for eBug to Slide
# Build with:  pyinstaller eBug_to_slide.spec
# Output:      dist/eBug_to_slide  (macOS) or  dist/eBug_to_slide.exe  (Windows)

block_cipher = None

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("Template/New Layout.pptx", "Template"),
    ],
    hiddenimports=[
        "browser_cookie3",
        "keyring.backends",
        "keyring.backends.macOS",
        "keyring.backends.Windows",
        "keyring.backends.SecretService",
        "keyring.backends.fail",
        "requests_ntlm",
        "lxml._elementpath",
        "PIL._tkinter_finder",
        "tkinter",
        "tkinter.ttk",
        "tkinter.scrolledtext",
        "tkinter.filedialog",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="eBug_to_slide",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,   # windowed — no terminal window on double-click
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
    name="eBug_to_slide",
)

# macOS .app bundle (onedir mode)
app = BUNDLE(
    coll,
    name="eBug_to_slide.app",
    icon=None,
    bundle_identifier="com.cyberlink.ebug-to-slide",
)

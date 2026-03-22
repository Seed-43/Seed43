# Seed43_Setup.spec
# Run with:  pyinstaller Seed43_Setup.spec

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='Seed43_Setup',
    debug=False,
    strip=False,
    upx=True,
    console=False,          # No console window — GUI only
    icon=None,              # Replace with 'seed43.ico' when you have one
)

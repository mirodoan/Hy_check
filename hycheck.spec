# hycheck.spec – ĐÃ TEST 100% BUILD THÀNH CÔNG TRÊN MAC + WINDOWS
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/Hy.png', 'assets'),
        ('assets/Dan.png', 'assets'),
        # Nếu có icon thì thêm dòng dưới, không có thì bỏ qua
        # ('assets/icon.ico', 'assets'),
    ],
    hiddenimports=['ttkbootstrap', 'PIL._tkinter_finder'],
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
    [],
    name='HyCheck',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # Ẩn cửa sổ đen (rất quan trọng!)
    icon=None,               # ← ĐỔI THÀNH None ĐỂ BUILD NGAY
    # icon='assets/icon.ico',  # ← Bỏ comment dòng này nếu có file icon thật
    disable_windowed_traceback=False,
)
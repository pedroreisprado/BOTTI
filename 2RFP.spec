# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['init.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/*', 'assets')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    icon='C:/2RFP/assets/2rfp-icon.ico', 
    version='1.0',                    
    description='Sugestão de novos produtos integrado ao sistema MIllennium da Linx',  
    name='2RFP - Sugestão de vendas',       
    company_name='2RFP Technology'
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='2RFP',
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
    onefile=True
)

# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['CWManageLabelMaker\\main.py'],
    pathex=['CWManageLabelMaker'],
    binaries=[],
    datas=[],
    hiddenimports=[],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CWManageLabelMaker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)


import shutil
shutil.copyfile('CWManageLabelMaker/config.toml.example', '{0}/config.toml'.format(DISTPATH))
shutil.copyfile('CWManageLabelMaker/template.label', '{0}/template.label'.format(DISTPATH))
shutil.copyfile('installers/DLS8Setup8.7.4.exe', '{0}/DLS8Setup8.7.4.exe'.format(DISTPATH))
shutil.copyfile('installers/nssm.exe', '{0}/nssm.exe'.format(DISTPATH))
# -*- mode: python ; coding: utf-8 -*-
block_cipher = None


a = Analysis(['src\\main.py'],
             pathex=['C:\\Users\\dipeshs\\PycharmProjects\\excelUtitlity'],
             binaries=[('C:\\Users\\dipeshs\\PycharmProjects\\excelUtitlity\\icons','icons'),
             ('C:\\Users\\dipeshs\\PycharmProjects\\excelUtitlity\\validations','validations')],
             datas=[],
             hiddenimports=['pyreadstat._readstat_writer',
             'pyreadstat.worker',],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='Utility',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='main')

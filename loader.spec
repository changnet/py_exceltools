# -*- mode: python -*-

block_cipher = None


a = Analysis(['loader.py'],
             pathex=['E:\\linux_share\\github\\py_exceltools'],
             binaries=[],
             datas=[],
             hiddenimports=['writer_lua','writer_xml','writer_json'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='loader',
          debug=False,
          strip=False,
          upx=True,
          console=True )

# -*- mode: python -*-

block_cipher = None


a = Analysis(['powerpointGenerator.py'],
             pathex=['D:\\Users\\Luke\\Documents\\16-17 sem 1\\capstone\\LCABiometric'],
             binaries=[],
             datas=[],
             hiddenimports=[],
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
          name='powerpointGenerator',
          debug=False,
          strip=False,
          upx=True,
          console=True )

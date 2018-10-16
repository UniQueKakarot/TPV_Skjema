# -*- mode: python -*-

block_cipher = None

extra_files = [('C:\\Users\\heiv085\\Documents\\Github\\TPV_Skjema\\TPV_Source\\configs', 'configs'),
			   ('C:\\Users\\heiv085\\Documents\\Github\\TPV_Skjema\\TPV_Source\\modules', 'modules'),
			   ('C:\\Users\\heiv085\\Documents\\Github\\TPV_Skjema\\TPV_Source\\tpv.ico', '.')]

a = Analysis(['main.py'],
             pathex=['C:\\Users\\heiv085\\Documents\\Github\\TPV_Skjema\\TPV_Source'],
             binaries=[],
             datas=extra_files,
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
          exclude_binaries=True,
          name='TPV',
          debug=False,
          strip=False,
          upx=True,
          console=False,
          icon='tpv.ico' )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='TPV 2.1.01')

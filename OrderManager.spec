# -*- mode: python -*-
a = Analysis(['S:\\Everyone\\Management Software\\OrderManager\\development\\OrderManager.py'],
             pathex=['S:\\Everyone\\Management Software\\OrderManager\\development'],
             hiddenimports=[],
             hookspath=None)

files_to_include = '''
interface.xrc,
Splash.png,
OrderManager.ico,
clock.png,
printer.png
'''

for file_to_include in files_to_include.split(','):
          file_to_include = file_to_include.strip()
          a.datas += [(file_to_include , file_to_include , 'DATA')]

pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=1,
          name=os.path.join('build\\pyi.win32\\OrderManager', 'OrderManager.exe'),
          debug=False,
          strip=None,
          upx=True,
          icon='OrderManager.ico',
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=None,
               upx=True,
               name=os.path.join('dist', 'OrderManager'))

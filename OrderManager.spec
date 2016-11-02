# -*- mode: python -*-

block_cipher = None


a = Analysis(['OrderManager.py'],
             pathex=['C:\\temp_build\\development', 'C:\\Python27\\libs\\', 'C:\\Python27\\Lib\\site-packages'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None,
             cipher=block_cipher)

files_to_include = '''
interface.xrc,
Splash.png,
OrderManager.ico,
clock.png,
printer.png,
VisualizeEtoForecast.xlsm
'''

for file_to_include in files_to_include.split(','):
          file_to_include = file_to_include.strip()
          a.datas += [(file_to_include , file_to_include , 'DATA')]


pyz = PYZ(a.pure,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='OrderManager.exe',
          debug=False,
          strip=None,
          upx=False,
          icon='OrderManager.ico',
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=None,
               upx=False,
               name='OrderManager')

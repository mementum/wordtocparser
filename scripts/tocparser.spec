# -*- mode: python -*-
# pathex=['d:\\Users\\lgsatellital\\Documents\\02-msys\\D00737824\\src\\tocparser'])
a = Analysis([os.path.join(HOMEPATH,'support\\_mountzlib.py'),
              os.path.join(HOMEPATH,'support\\useUnicode.py'), 'src/tocparser.pyw'],
             pathex=None)
pyz = PYZ(a.pure)

exe = EXE(pyz,
          a.scripts,
          exclude_binaries=1,
          name=os.path.join('build\\pyi.win32\\tocparser', 'tocparser.exe'),
          debug=False,
          strip=True,
          upx=False,
          console=False)

coll = COLLECT( exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=False,
               name=os.path.join('dist', 'tocparser'))
app = BUNDLE(coll,
             name=os.path.join('dist', 'tocparser.app'))

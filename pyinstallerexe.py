""" run pyinstaller via program to add lengthy command"""

import PyInstaller.__main__

PyInstaller.__main__.run([
    '--distpath ./dist',
    '--onefile',
    'ROVCleaver V16-3dev.py'
])

exit()

'--hidden-import openpyxl.cell._writer',
# '--debug=imports',

# conda activate general2
# cd '/Users/Denise/Library/CloudStorage/Dropbox/Postcard Files/PythonProgs/ROVCleaver_on_Dropbox'
# PyInstaller --distpath . --onefile 'ROVCleaver V16-3dev.py'

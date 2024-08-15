from setuptools import setup

APP = ['excel.py']
DATA_FILES = [
    ('data', ['data/Plan et Rapport d\'audit certification V32.xlsm']),
]

OPTIONS = {
    'argv_emulation': True,
    'packages': ['xlwings'],
    'includes': ['xlwings', 'PyQt5'],
    'resources': ['data'],
    'plist': 'Info.plist',
    'excludes': ['PyQt6'],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)

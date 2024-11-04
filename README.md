poetry run pyinstaller --name Proneo --onefile --add-data "data/Plan et Rapport d'audit certification V33.xlsm;data" --hidden-import xlwings --hidden-import PyQt5 --windowed excel.py

poetry run pyinstaller --name Proneo --onefile --add-data "data/Plan et Rapport d'audit certification V33.xlsm;data" --hidden-import xlwings --hidden-import PyQt5 --console excel.py

git reset --hard origin/main

sudo poetry run python setup.py py2app

@echo off
pip install -r requirements.txt
pyinstaller --onefile --windowed --name TTN_Generator --add-data "template_ttn.xlsx;." ttn_generator.py
echo.
echo Done: dist\TTN_Generator.exe
pause

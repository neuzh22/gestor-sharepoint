@echo off
rem Genera dist\SP-Migrar.exe (un solo archivo, no requiere Python en el equipo destino).
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install --upgrade -r requirements.txt pyinstaller || exit /b 1
python -m PyInstaller --noconfirm --clean --onefile --console ^
    --name SP-Migrar ^
    --collect-data certifi ^
    sp_migrar.py || exit /b 1
echo.
echo Listo: dist\SP-Migrar.exe

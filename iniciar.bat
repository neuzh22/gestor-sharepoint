@echo off
chcp 65001 >nul
cd /d "%~dp0"
python --version >nul 2>&1 || (echo No se encontro Python. Instala Python 3.10+ desde python.org & pause & exit /b 1)
python -c "import msal, requests" >nul 2>&1 || python -m pip install -r requirements.txt
python sp_migrar.py %*
pause

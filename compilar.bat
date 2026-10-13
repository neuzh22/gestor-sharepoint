@echo off
rem Genera dist\SP-Migrar.exe (un solo archivo, no requiere Python en el equipo destino).
rem Si existe ffmpeg\ffmpeg.exe lo copia junto al .exe para la compresión de imágenes.
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install --upgrade -r requirements.txt pyinstaller || exit /b 1
python -m PyInstaller --noconfirm --clean --onefile --console ^
    --name SP-Migrar ^
    --collect-data certifi ^
    sp_migrar.py || exit /b 1
if exist ffmpeg\ffmpeg.exe (
    copy /y ffmpeg\ffmpeg.exe dist\ffmpeg.exe >nul
    echo FFmpeg copiado a dist\ffmpeg.exe
) else (
    echo AVISO: no hay ffmpeg\ffmpeg.exe; la compresion de imagenes no estara disponible.
    echo        Descargalo de https://www.gyan.dev/ffmpeg/builds/ ^(release essentials^) y ponlo en ffmpeg\ffmpeg.exe
)
echo.
echo Listo: copia el contenido de la carpeta dist\ ^(SP-Migrar.exe y ffmpeg.exe^) al equipo donde lo usaras.

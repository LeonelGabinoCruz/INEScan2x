@echo off
rem Regenera el .exe y el instalador despues de cambiar escaneo_ine.py
cd /d "%~dp0"
python -m PyInstaller --noconfirm --onefile --windowed --name INEScan2x --icon "%~dp0icono.ico" --add-data "%~dp0icono.ico;." --hidden-import win32timezone --hidden-import pythoncom --hidden-import win32com.client --distpath "%~dp0INSTALADOR" --workpath "%TEMP%\inescan_build" --specpath "%TEMP%\inescan_build" escaneo_ine.py || goto :fin
"%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" "%~dp0INSTALADOR\instalador.iss"
echo.
echo Listo: INSTALADOR\salida\Instalar_INEScan2x.exe
:fin
pause

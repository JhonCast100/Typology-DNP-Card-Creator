@echo off
chcp 65001 >nul
echo Instalando dependencias si no están disponibles...
python -m pip install pyinstaller openpyxl pillow reportlab pandas

echo.
echo Limpiando compilaciones anteriores...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul

echo.
echo Generando ejecutable...
python -m PyInstaller build.spec --distpath=./dist --workpath=./build

echo.
echo Configurando carpetas de distribución...
python setup_dist.py

echo.
echo.
echo ====================================
echo ✓ Construccion completada!
echo.
echo El ejecutable se encuentra en: dist\DNPCardCreator\DNPCardCreator.exe
echo.
echo Carpetas configuradas:
echo - dist\DNPCardCreator\DNPCardCreator.exe
echo - dist\DNPCardCreator\Data\
echo - dist\DNPCardCreator\Output\
echo ====================================
pause

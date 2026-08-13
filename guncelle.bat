@echo off
setlocal

cd /d "%~dp0"

echo.
echo ==========================================
echo       PS26 EXCEL VERI GUNCELLEME
echo ==========================================
echo.

if not exist "data.xlsx" (
    echo HATA: data.xlsx bulunamadi!
    echo.
    pause
    exit /b 1
)

if not exist "convert.py" (
    echo HATA: convert.py bulunamadi!
    echo.
    pause
    exit /b 1
)

echo Excel verileri donusturuluyor...
echo.

python convert.py

if errorlevel 1 (
    echo.
    echo ==========================================
    echo HATA: Donusturme basarisiz!
    echo ==========================================
    echo.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo BASARILI!
echo ==========================================
echo.
echo JSON dosyalari guncellendi.
echo.

pause
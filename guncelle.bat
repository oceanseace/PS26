@echo off
setlocal

cd /d "%~dp0"

echo.
echo ==========================================
echo       PS26 EXCEL GUNCELLEME
echo ==========================================
echo.

REM ------------------------------------------
REM data.xlsx kontrolu
REM ------------------------------------------

if not exist "data.xlsx" (
    echo HATA: data.xlsx bulunamadi!
    echo.
    pause
    exit /b 1
)

REM ------------------------------------------
REM convert.py kontrolu
REM ------------------------------------------

if not exist "convert.py" (
    echo HATA: convert.py bulunamadi!
    echo.
    pause
    exit /b 1
)

REM ------------------------------------------
REM EXCEL -> JSON
REM ------------------------------------------

echo [1/4] Excel verileri donusturuluyor...
echo.

python .\convert.py

if errorlevel 1 (
    echo.
    echo HATA: Excel donusturme basarisiz!
    echo.
    pause
    exit /b 1
)

echo.
echo [OK] JSON dosyalari guncellendi.
echo.


REM ------------------------------------------
REM GIT ADD
REM ------------------------------------------

echo [2/4] Git degisiklikleri ekleniyor...
echo.

git add index.html data

if errorlevel 1 (
    echo.
    echo HATA: git add basarisiz!
    echo.
    pause
    exit /b 1
)

echo.
echo [OK] Dosyalar Git'e eklendi.
echo.


REM ------------------------------------------
REM GIT COMMIT
REM ------------------------------------------

echo [3/4] Commit yapiliyor...
echo.

git commit -m "update"

if errorlevel 1 (
    echo.
    echo UYARI: Commit yapilacak yeni degisiklik olmayabilir.
    echo.
)


REM ------------------------------------------
REM GIT PUSH
REM ------------------------------------------

echo [4/4] GitHub'a gonderiliyor...
echo.

git push

if errorlevel 1 (
    echo.
    echo ==========================================
    echo HATA: GitHub push basarisiz!
    echo ==========================================
    echo.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo          BASARIYLA TAMAMLANDI
echo ==========================================
echo.
echo JSON verileri guncellendi.
echo GitHub commit yapildi.
echo GitHub push yapildi.
echo.
echo GitHub Pages birkac dakika icinde guncellenir.
echo.

pause
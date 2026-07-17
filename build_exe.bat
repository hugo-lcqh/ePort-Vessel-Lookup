@echo off
REM ============================================================
REM  Dong goi eport_vessel_lookup.py thanh file .exe
REM  Chay file nay (double-click) tren may da cai Python
REM ============================================================
chcp 65001 >nul
echo.
echo ============================================================
echo   DONG GOI ePort Vessel Lookup -^> file .exe
echo ============================================================
echo.

REM Buoc 1: Cai cac thu vien can thiet
echo [1/3] Dang cai thu vien...
python -m pip install --upgrade selenium webdriver-manager colorama tabulate openpyxl pyinstaller
echo.

REM Buoc 2: Dong goi bang PyInstaller (goi qua python -m de tranh loi PATH)
echo [2/3] Dang dong goi (co the mat 1-3 phut)...
python -m PyInstaller --onefile --console ^
  --name "ePortVesselLookup" ^
  --collect-all webdriver_manager ^
  --collect-all selenium ^
  eport_vessel_lookup.py
echo.

REM Buoc 3: Hoan tat
echo [3/3] Hoan tat!
echo.
if exist "dist\ePortVesselLookup.exe" (
    echo ============================================================
    echo   THANH CONG!
    echo   File .exe nam tai: dist\ePortVesselLookup.exe
    echo   Ban co the copy file nay sang may khac de chay.
    echo ============================================================
) else (
    echo ============================================================
    echo   CO LOI XAY RA. Kiem tra thong bao loi o tren.
    echo ============================================================
)
echo.
pause

@echo off
title Build TF Readings EXE
echo ====================================================
echo             Building TF Readings EXE
echo ====================================================
echo.

where py >nul 2>nul
if %ERRORLEVEL% equ 0 (
    set PY_CMD=py
) else (
    set PY_CMD=python
)

echo Using Python launcher: %PY_CMD%
echo.
echo Installing / checking required packages...
%PY_CMD% -m pip install --upgrade pip
%PY_CMD% -m pip install pandas openpyxl tkinterdnd2 pyinstaller

echo.
echo Compiling to standalone EXE...
%PY_CMD% -m PyInstaller --noconsole --onefile --collect-data tkinterdnd2 dat_to_excel_tool.py

echo.
if exist "dist\dat_to_excel_tool.exe" (
    echo ====================================================
    echo SUCCESS: Standalone EXE built successfully!
    echo Location: dist\dat_to_excel_tool.exe
    echo ====================================================
) else (
    echo ====================================================
    echo BUILD FAILED. Check errors above.
    echo ====================================================
)
pause

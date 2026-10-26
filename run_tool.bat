@echo off
title TF Readings
echo Starting TF Readings...
where py >nul 2>nul
if %ERRORLEVEL% equ 0 (
    py dat_to_excel_tool.py
) else (
    python dat_to_excel_tool.py
)
if %ERRORLEVEL% neq 0 (
    pause
)

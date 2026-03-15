@echo off
title FPL Auto-Refresh — Nippon Paint QA Cup
color 0A

echo.
echo  =====================================================
echo   FPL Auto-Refresh — Nippon Paint QA Cup
echo   Sedang mengambil data terbaru dari internet...
echo  =====================================================
echo.

:: Cari Python — coba py launcher dulu, lalu python, lalu python3
set PYTHON=
where py >nul 2>nul && set PYTHON=py
if "%PYTHON%"=="" (
    where python >nul 2>nul && set PYTHON=python
)
if "%PYTHON%"=="" (
    where python3 >nul 2>nul && set PYTHON=python3
)

if "%PYTHON%"=="" (
    echo  [ERROR] Python tidak ditemukan di komputer ini.
    echo.
    echo  Silakan download Python di: https://www.python.org/downloads/
    echo  Centang "Add Python to PATH" saat instalasi.
    echo.
    pause
    exit /b 1
)

:: Cek dan install openpyxl kalau belum ada
%PYTHON% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo  [INFO] Menginstall openpyxl (hanya sekali)...
    %PYTHON% -m pip install openpyxl --quiet
    if errorlevel 1 (
        echo  [ERROR] Gagal install openpyxl. Coba jalankan sebagai Administrator.
        pause
        exit /b 1
    )
)

:: Jalankan script
%PYTHON% "%~dp0fpl_refresh.py"

if errorlevel 1 (
    echo.
    echo  [ERROR] Script gagal dijalankan.
    pause
)

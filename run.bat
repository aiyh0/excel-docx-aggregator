@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

REM Folder you launched from (so auto-detect scans there)
set "CALLER_DIR=%CD%"
pushd "%~dp0"

set "SCRIPT=%~dp0excel_to_docx_aggregate.py"
set "VENV=%~dp0venv"
set "VENV_PY=%VENV%\Scripts\python.exe"
set "REQ=%~dp0requirements.txt"
set "REQHASH=%VENV%\requirements.hash"
set "TMPHASH=%TEMP%\reqhash_%RANDOM%.txt"

REM --- Find a base Python to create venv (prefer py launcher) ---
set "BASEPY="
where py >nul 2>&1 && set "BASEPY=py -3"
if not defined BASEPY (
  where python >nul 2>&1 && set "BASEPY=python"
)
if not defined BASEPY (
  echo [Error] Python 3 not found on PATH. Please install Python 3 and re-run.
  goto :end
)

REM --- Create venv on first run ---
if not exist "%VENV_PY%" (
  echo [Setup] Creating virtual environment...
  %BASEPY% -m venv "%VENV%" || ( echo [Error] Failed to create venv & goto :end )
)

REM --- Ensure requirements.txt exists (create a default if missing) ---
if not exist "%REQ%" (
  echo [Setup] Creating default requirements.txt...
  >"%REQ%" echo pandas^>=2.0
  >>"%REQ%" echo openpyxl^>=3.1
  >>"%REQ%" echo python-docx^>=1.1
  >>"%REQ%" echo rapidfuzz^>=3.0
)

REM --- Check if requirements changed (use certutil if available) ---
set "NEED_INSTALL=1"
where certutil >nul 2>&1
if %errorlevel%==0 (
  certutil -hashfile "%REQ%" MD5 > "%TMPHASH%" 2>nul
  if exist "%REQHASH%" (
     fc /b "%TMPHASH%" "%REQHASH%" >nul 2>&1 && set "NEED_INSTALL=0"
  )
) else (
  echo [Info] certutil not found; will install requirements without hash check.
)

if "!NEED_INSTALL!"=="1" (
  echo [Setup] Installing/updating requirements into venv...
  "%VENV_PY%" -m pip install --upgrade pip || ( echo [Error] pip upgrade failed & goto :end )
  "%VENV_PY%" -m pip install -r "%REQ%"   || ( echo [Error] package install failed & goto :end )
  if exist "%TMPHASH%" move /y "%TMPHASH%" "%REQHASH%" >nul
) else (
  del "%TMPHASH%" >nul 2>&1
)

REM --- Run: file, folder, or auto-detect in caller dir ---
if "%~1"=="" (
  echo [Run] Auto-select newest Excel in: %CALLER_DIR%
  "%VENV_PY%" "%SCRIPT%" --dir "%CALLER_DIR%"
) else (
  if exist "%~1\" (
    echo [Run] Auto-select newest Excel in folder: %~1
    "%VENV_PY%" "%SCRIPT%" --dir "%~1"
  ) else (
    echo [Run] Using file: %~1
    "%VENV_PY%" "%SCRIPT%" "%~1"
  )
)

if errorlevel 1 (
  echo.
  echo [!] The script reported an error. See messages above.
  pause
)

: end
popd
endlocal

@echo off
setlocal enabledelayedexpansion

REM ===== Paths =====
set "ACAD=C:\Program Files\Autodesk\AutoCAD 2022\accoreconsole.exe"
set "FOLDER=C:\Users\admin\Documents\AUTOCAD_WEBAPP\DXF"
set "SCR=C:\Users\admin\Documents\AUTOCAD_WEBAPP\run_visjsonnet.scr"

REM Your python exe (recommended: explicit)
set "PY=C:\Users\admin\.pyenv\pyenv-win\versions\3.7.9\python.exe"

REM Your uploader script path
set "UPLOAD_PY=C:\Users\admin\Documents\AUTOCAD_WEBAPP\Backend\upload_json_to_sheet_export.py"

REM Where AutoCAD writes json
set "EXPORT_DIR=C:\Users\admin\Documents\AUTOCAD_WEBAPP\EXPORTS"
set "JSON1=%EXPORT_DIR%\vis_export_visibility.json"
set "JSON2=%EXPORT_DIR%\vis_export_all.json"

REM ===== Get newest DWG =====
for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command ^
  "(Get-ChildItem -Path '%FOLDER%' -Filter *.dwg | Sort-Object LastWriteTime -Desc | Select-Object -First 1).FullName"`) do set "DWG=%%F"

if not defined DWG (
  echo ❌ No DWG found in: %FOLDER%
  exit /b 1
)

echo ✅ Using DWG: %DWG%

REM ===== Run AutoCAD Core Console =====
"%ACAD%" /i "%DWG%" /s "%SCR%"
if errorlevel 1 (
  echo ❌ AutoCAD accoreconsole failed with errorlevel %errorlevel%
  exit /b %errorlevel%
)

echo ✅ AutoCAD done. Waiting for JSON exports...

REM ===== Wait for JSONs to exist (max 60 seconds) =====
set /a tries=0
:waitloop
set /a tries+=1

if exist "%JSON1%" if exist "%JSON2%" goto gotjson

if %tries% GEQ 60 (
  echo ❌ Timed out waiting for JSON files:
  echo    %JSON1%
  echo    %JSON2%
  exit /b 2
)

timeout /t 1 >nul
goto waitloop

:gotjson
echo ✅ JSON files found.
echo    %JSON1%
echo    %JSON2%

REM ===== Upload to Google Sheets =====
echo ⬆ Uploading to Google Sheets...
"%PY%" "%UPLOAD_PY%"
if errorlevel 1 (
  echo ❌ Upload script failed with errorlevel %errorlevel%
  exit /b %errorlevel%
)

echo 🎉 Done: AutoCAD export + Sheets upload completed successfully.
exit /b 0

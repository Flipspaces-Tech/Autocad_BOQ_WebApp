@echo off
setlocal

set "ACAD=C:\Program Files\Autodesk\AutoCAD 2022\accoreconsole.exe"
set "FOLDER=C:\Users\admin\Documents\AUTOCAD_WEBAPP\DXF"
set "SCR=C:\Users\admin\Documents\AUTOCAD_WEBAPP\run_visjsonnet.scr"

REM Get newest .dwg by modified date using PowerShell (returns full path)
for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command ^
  "(Get-ChildItem -Path '%FOLDER%' -Filter *.dwg | Sort-Object LastWriteTime -Desc | Select-Object -First 1).FullName"`) do set "DWG=%%F"

if not defined DWG (
  echo No DWG found in: %FOLDER%
  exit /b 1
)

echo Using DWG: %DWG%
"%ACAD%" /i "%DWG%" /s "%SCR%"
exit /b %errorlevel%

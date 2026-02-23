@echo off
set ACAD="C:\Program Files\Autodesk\AutoCAD 2022\accoreconsole.exe"
set DWG="C:\Users\admin\Documents\AUTOCAD_WEBAPP\DXF\BLOCK (FURNITURE) MM (1).dwg"
set SCR="C:\Users\admin\Documents\AUTOCAD_WEBAPP\run_visjsonnet.scr"

%ACAD% /i %DWG% /s %SCR%

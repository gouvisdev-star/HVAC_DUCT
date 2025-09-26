@echo off
echo Setting up HVAC_DUCT Auto-load system...

REM Get AutoCAD support folder
set "ACADSUPPORT=%APPDATA%\Autodesk\AutoCAD 2024\R24.3\enu\Support"
if not exist "%ACADSUPPORT%" (
    set "ACADSUPPORT=%APPDATA%\Autodesk\AutoCAD 2023\R23.1\enu\Support"
)
if not exist "%ACADSUPPORT%" (
    set "ACADSUPPORT=%APPDATA%\Autodesk\AutoCAD 2022\R22.0\enu\Support"
)
if not exist "%ACADSUPPORT%" (
    set "ACADSUPPORT=%APPDATA%\Autodesk\AutoCAD 2021\R24.0\enu\Support"
)

echo AutoCAD Support Folder: %ACADSUPPORT%

REM Copy LISP files to AutoCAD support folder
copy "HVAC_Simple.lsp" "%ACADSUPPORT%\HVAC_Simple.lsp"
if %errorlevel%==0 (
    echo LISP file copied successfully!
) else (
    echo Failed to copy LISP file. Please copy manually.
)

REM Create acad.lsp if it doesn't exist
if not exist "%ACADSUPPORT%\acad.lsp" (
    echo (load "HVAC_Simple.lsp") > "%ACADSUPPORT%\acad.lsp"
    echo Created acad.lsp
) else (
    echo acad.lsp already exists
)

REM Add to ACADLSPASDOC
echo.
echo To complete setup:
echo 1. Open AutoCAD
echo 2. Type: ACADLSPASDOC
echo 3. Add: %ACADSUPPORT%\HVAC_Simple.lsp
echo 4. Restart AutoCAD
echo.
echo Or simply type: HVAC in AutoCAD command line

echo.
echo Setup complete!
pause

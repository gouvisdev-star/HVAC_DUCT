@echo off
echo Setting up HVAC_DUCT Auto-load...

REM Get AutoCAD support folder
set "ACADSUPPORT=%APPDATA%\Autodesk\AutoCAD 2024\R24.3\enu\Support"
if not exist "%ACADSUPPORT%" (
    set "ACADSUPPORT=%APPDATA%\Autodesk\AutoCAD 2023\R23.1\enu\Support"
)
if not exist "%ACADSUPPORT%" (
    set "ACADSUPPORT=%APPDATA%\Autodesk\AutoCAD 2022\R22.0\enu\Support"
)

echo AutoCAD Support Folder: %ACADSUPPORT%

REM Copy LISP file to AutoCAD support folder
copy "acad.lsp" "%ACADSUPPORT%\acad.lsp"
if %errorlevel%==0 (
    echo LISP file copied successfully!
) else (
    echo Failed to copy LISP file. Please copy manually.
)

REM Add to ACADLSPASDOC system variable
echo.
echo To enable auto-load, add this to AutoCAD:
echo 1. Type: ACADLSPASDOC
echo 2. Add: %ACADSUPPORT%\acad.lsp
echo 3. Or add: HVACSTARTUP to startup sequence

echo.
echo Setup complete!
pause

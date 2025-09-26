HVAC_DUCT Auto-Load Setup Instructions
=====================================

QUICK SETUP (Recommended):
1. Run: setup_hvac.bat
2. Open AutoCAD
3. Type: HVAC
4. Done!

MANUAL SETUP:
1. Copy HVAC_Simple.lsp to AutoCAD support folder:
   %APPDATA%\Autodesk\AutoCAD 2024\R24.3\enu\Support\
   
2. Add to ACADLSPASDOC system variable:
   - Type: ACADLSPASDOC
   - Add: Path to HVAC_Simple.lsp
   
3. Restart AutoCAD

USAGE:
- HVAC: Load DLL + Enable auto-cleanup + Load ticks
- AUTO: Load ticks manually
- CLEANUPTICKS: Clean up deleted objects

AUTOMATIC FEATURES:
- Ticks load automatically when opening files
- Deleted objects are automatically cleaned from tick file
- No manual commands needed after setup

FILES:
- HVAC_Simple.lsp: Simple auto-load script
- HVAC_DUCT_Auto.lsp: Full-featured script
- setup_hvac.bat: Automatic setup script

TROUBLESHOOTING:
- If DLL not found: Check path in HVAC_Simple.lsp
- If not auto-loading: Check ACADLSPASDOC setting
- If errors: Type HVAC-HELP for help commands

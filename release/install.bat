@ECHO OFF
SETLOCAL

cd /d "%~dp0"

echo Installing PdbToExcel and dependencies...
echo.

pip install --find-links packages pnfl-pdbtoexcel
if %ERRORLEVEL% NEQ 0 goto :error

echo.
echo Installation complete!
echo.
echo Next steps:
echo   1. Edit pdb_to_excel.ini with your team name and PNFL path
echo   2. Run PdbToExcel.bat
echo.
pause
exit /b 0

:error
echo.
echo ERROR: Installation failed. See above for details.
pause
exit /b 1

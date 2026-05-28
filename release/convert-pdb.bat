@ECHO OFF
cd /d "%~dp0"

:: ===================================================================
:: Edit the paths below to match your setup, then double-click to run.
:: File extension matters! Use .xlsm for macro-enabled sorting.
:: ===================================================================

SET "PDB_FILE=C:\PATH\TO\STATS.pdb"
SET "OUTPUT_FILE=C:\PATH\TO\OUTPUT.xlsm"
SET "DEF1_PLN=C:\PATH\TO\DEFENSE1.pln"
SET "OFF1_PLN=C:\PATH\TO\OFFENSE1.pln"
SET "DEF2_PLN=C:\PATH\TO\DEFENSE2.pln"
SET "OFF2_PLN=C:\PATH\TO\OFFENSE2.pln"

pnfl convert-pdb "%PDB_FILE%" "%OUTPUT_FILE%" -d "%DEF1_PLN%" -o "%OFF1_PLN%" -d2 "%DEF2_PLN%" -o2 "%OFF2_PLN%"

pause

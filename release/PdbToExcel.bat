@ECHO OFF
cd /d "%~dp0"

:: ===================================================================
:: Edit the paths below to match your setup, then double-click to run.
:: File extension matters! Use .xlsm for macro-enabled sorting.
:: ===================================================================

SET PDB_FILE=C:\PATH\TO\STATS.pdb
SET OUTPUT_FILE=C:\PATH\TO\OUTPUT.xlsm
SET DEFENSE_PLN=C:\PATH\TO\DEFENSE.pln
SET OFFENSE_PLN=C:\PATH\TO\OFFENSE.pln

python -m pnfl_pdbtoexcel "%PDB_FILE%" "%OUTPUT_FILE%" -d "%DEFENSE_PLN%" -o "%OFFENSE_PLN%"

pause

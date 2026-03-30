@ECHO OFF

SET "PYTHONPATH=%~dp0..\src"

:: Sample batch file for running the installed package entry module from the
:: project checkout.

:: File extension matters! Use .xlsm if you want double-click sorting.

py -m pnfl_pdbtoexcel -d "C:\DEF\DGP.pln" -o "C:\OFF\OGP.pln" "C:\PATH\TO\PDB.pdb" "C:\PATH\TO\XLSM.xlsm"

pause

:: Launch Excel
start "" "C:\PATH\TO\XLSM.xlsm"

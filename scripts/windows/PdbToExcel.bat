@ECHO OFF

:: Sample batch file for a distributed folder where this .bat sits next to
:: PdbToExcel.py and PdbToExcel.ini.

:: File extension matters! Use .xlsm if you want double-click sorting.

python PdbToExcel.py -d "C:\DEF\DGP.pln" -o "C:\OFF\OGP.pln" "C:\PATH\TO\PDB.pdb" "C:\PATH\TO\XLSM.xlsm"

pause

:: Launch Excel
start "" "C:\PATH\TO\XLSM.xlsm"

@ECHO OFF
SETLOCAL ENABLEDELAYEDEXPANSION

cd /d "%~dp0"

:: NOTE: Must use file extension .xlsm to enable double-click sorting

:: 1 = use one hardcoded O/D game plan for all weeks
:: 0 = use week-specific O/D game plans (47W1, 47W2, ...)
SET USE_HARDCODED_PLANS=1

SET SEASON=2048
SET "TEAM_DIR=E:\SIERRA\FbPro98\PNFL\%SEASON%\Plans\Denver (Brian)"
SET "GAMELOG_DIR=E:\PNFL\League - Game Logs"

:: Default game plans
SET "DEFAULT_OGP=%TEAM_DIR%\DEN-OGP1.pln"
SET "DEFAULT_DGP=%TEAM_DIR%\DEN-DGP1.pln"

:: Fixed game plans (used when USE_HARDCODED_PLANS=1)
SET "HARDCODED_OGP=%DEFAULT_OGP%"
SET "HARDCODED_DGP=%DEFAULT_DGP%"
REM SET "HARDCODED_OGP=E:\SIERRA\FbPro98\PNFL\2047\Plans\Denver (Brian)\47W9\DEN-OGP1.pln"
REM SET "HARDCODED_DGP=E:\SIERRA\FbPro98\PNFL\2047\Plans\Denver (Brian)\47W9\DEN-DGP1.pln"

:: Weekly spreadsheet (re-create past weeks of current season with latest game plans)

FOR /L %%W IN (1,1,19) DO (
	SET "WEEK=!SEASON:~-2!W%%W"
	IF "%USE_HARDCODED_PLANS%"=="1" (
        :: Hardcoded game plan for all weeks
        SET "OGP=%HARDCODED_OGP%"
        SET "DGP=%HARDCODED_DGP%"
	) ELSE (
        :: Week-specific game plans, e.g. ...\47W4\DEN-DGP1.pln
        SET "OGP=%TEAM_DIR%\!WEEK!\DEN-OGP1.pln"		
        SET "DGP=%TEAM_DIR%\!WEEK!\DEN-DGP1.pln"
	)
    SET "PDB=%GAMELOG_DIR%\!SEASON!\!WEEK!.pdb"
    SET "XLSM=%GAMELOG_DIR%\!SEASON!\!WEEK!.xlsm"
	
    IF EXIST "!PDB!" (
        pnfl convert-pdb "!PDB!" "!XLSM!" -o "!OGP!" -d "!DGP!" --skip-calcs
    )
)

:: Yearly/Multi-year spreadsheets (re-create with latest data and game plans)
FOR %%F IN ("%GAMELOG_DIR%\*.PDB") DO (
    SET "PDB=%%~fF"
    SET "XLSM=%%~dpnF.xlsm"

	pnfl convert-pdb "!PDB!" "!XLSM!" -o "%HARDCODED_OGP%" -d "%HARDCODED_DGP%"
)

ENDLOCAL

PAUSE

@ECHO OFF
:: ENABLEDELAYEDEXPANSION lets us use !VAR! inside FOR/IF blocks to read
:: values set in the same iteration. %VAR% expands at parse time (once,
:: before the block runs) and would give stale or empty values.
SETLOCAL ENABLEDELAYEDEXPANSION

:: NOTE: Must use file extension .xlsm to enable double-click sorting

SET SEASON=2048

SET "FBPRO_DIR=E:\SIERRA\FbPro98"
SET "TEAM_DIR=%FBPRO_DIR%\PNFL\%SEASON%\Plans\Denver (Brian)"
SET "GAMELOG_DIR=E:\PNFL\League - Game Logs"

:: 1 = use week-specific O/D game plans for weekly worksheets (47W1, 47W2, ...)
SET WEEKLY_PLANS=1
SET TWO_OFFENSIVE_PLANS=1
SET TWO_DEFENSIVE_PLANS=1

:: Game plan names
SET OFF_PLAN1=DEN-OGP1.pln
SET OFF_PLAN2=DEN-OGP2.pln
SET DEF_PLAN1=DEN-DGP1.pln
SET DEF_PLAN2=DEN-DGP2.pln

:: Default game plans
SET "DEFAULT_OGP1=%FBPRO_DIR%\%OFF_PLAN1%"
SET "DEFAULT_OGP2=%FBPRO_DIR%\%OFF_PLAN2%"
SET "DEFAULT_DGP1=%FBPRO_DIR%\%DEF_PLAN1%"
SET "DEFAULT_DGP2=%FBPRO_DIR%\%DEF_PLAN2%"

:: Hardcoded game plans
SET "SELECTED_OGP1=%DEFAULT_OGP1%"
SET "SELECTED_OGP2=%DEFAULT_OGP2%"
SET "SELECTED_DGP1=%DEFAULT_DGP1%"
SET "SELECTED_DGP2=%DEFAULT_DGP2%"
REM SET "SELECTED_OGP1=E:\SIERRA\FbPro98\PNFL\2047\Plans\Denver (Brian)\47W9\DEN-OGP1.pln"
REM SET "SELECTED_OGP2=E:\SIERRA\FbPro98\PNFL\2047\Plans\Denver (Brian)\47W9\DEN-DGP1.pln"

:: Weekly spreadsheet (re-create past weeks of current season with latest game plans)
FOR /L %%W IN (1,1,19) DO (
	SET "WEEK=!SEASON:~-2!W%%W"
	IF "%WEEKLY_PLANS%"=="1" (
        :: Week-specific game plans, e.g. ...\47W4\DEN-DGP1.pln
        SET "OGP1=%TEAM_DIR%\!WEEK!\%OFF_PLAN1%"
    	IF "%TWO_OFFENSIVE_PLANS%"=="1" (
            SET "OGP2=%TEAM_DIR%\!WEEK!\%OFF_PLAN2%"
        )
        SET "DGP1=%TEAM_DIR%\!WEEK!\%DEF_PLAN1%"
    	IF "%TWO_DEFENSIVE_PLANS%"=="1" (
            SET "DGP2=%TEAM_DIR%\!WEEK!\%DEF_PLAN2%"
        )
	) ELSE (
        :: Hardcoded game plan for all weeks
        SET "OGP1=%SELECTED_OGP1%"
        IF "%TWO_OFFENSIVE_PLANS%"=="1" (
            SET "OGP2=%SELECTED_OGP2%"
        )
        SET "DGP1=%SELECTED_DGP1%"
        IF "%TWO_DEFENSIVE_PLANS%"=="1" (
            SET "DGP2=%SELECTED_DGP2%"
        )
	)
    SET "PDB=%GAMELOG_DIR%\!SEASON!\!WEEK!.pdb"
    SET "XLSM=%GAMELOG_DIR%\!SEASON!\!WEEK!.xlsm"

    IF EXIST "!PDB!" (
        SET "EXTRA_PLANS="
        IF DEFINED OGP2 SET "EXTRA_PLANS=!EXTRA_PLANS! -o2 "!OGP2!""
        IF DEFINED DGP2 SET "EXTRA_PLANS=!EXTRA_PLANS! -d2 "!DGP2!""
        pnfl convert-pdb "!PDB!" "!XLSM!" -o "!OGP1!" -d "!DGP1!" !EXTRA_PLANS! --skip-calcs
    )
)

:: Yearly/Multi-year spreadsheets (re-create with latest data and game plans).
SET "YEARLY_EXTRA_PLANS="
IF "%TWO_OFFENSIVE_PLANS%"=="1" SET "YEARLY_EXTRA_PLANS=%YEARLY_EXTRA_PLANS% -o2 "%SELECTED_OGP2%""
IF "%TWO_DEFENSIVE_PLANS%"=="1" SET "YEARLY_EXTRA_PLANS=%YEARLY_EXTRA_PLANS% -d2 "%SELECTED_DGP2%""

FOR %%F IN ("%GAMELOG_DIR%\*.PDB") DO (
    pnfl convert-pdb "%%~fF" "%%~dpnF.xlsm" -o "%SELECTED_OGP1%" -d "%SELECTED_DGP1%" %YEARLY_EXTRA_PLANS%
)

:: Example individual calls
:: SEAONAL
::   pnfl convert-pdb "E:\PNFL\League - Game Logs\2048.pdb" "E:\PNFL\League - Game Logs\2048.xlsm" -o "E:\SIERRA\FbPro98\DEN-OGP1.pln" -d "E:\SIERRA\FbPro98\DEN-DGP1.pln"
::   pnfl convert-pdb "E:\PNFL\League - Game Logs\2048.pdb" "E:\PNFL\League - Game Logs\2048.xlsm" -o "E:\SIERRA\FbPro98\DEN-OGP1.pln" -d "E:\SIERRA\FbPro98\DEN-DGP1.pln" -o2 "E:\SIERRA\FbPro98\DEN-OGP2.pln" -d2 "E:\SIERRA\FbPro98\DEN-DGP2.pln"
:: WEEKLY
::   pnfl convert-pdb "E:\PNFL\League - Game Logs\2048\48W12.pdb" "E:\PNFL\League - Game Logs\2048\48W12.xlsm" -o "E:\SIERRA\FbPro98\PNFL\2048\Plans\Denver (Brian)\48W12\DEN-OGP1.pln" -d "E:\SIERRA\FbPro98\PNFL\2048\Plans\Denver (Brian)\48W12\DEN-DGP1.pln" --skip-calcs
::   pnfl convert-pdb "E:\PNFL\League - Game Logs\2048\48W12.pdb" "E:\PNFL\League - Game Logs\2048\48W12.xlsm" -o "E:\SIERRA\FbPro98\PNFL\2048\Plans\Denver (Brian)\48W12\DEN-OGP1.pln" -d "E:\SIERRA\FbPro98\PNFL\2048\Plans\Denver (Brian)\48W12\DEN-DGP1.pln" -o2 "E:\SIERRA\FbPro98\PNFL\2048\Plans\Denver (Brian)\48W12\DEN-OGP2.pln" -d2 "E:\SIERRA\FbPro98\PNFL\2048\Plans\Denver (Brian)\48W12\DEN-DGP2.pln" --skip-calcs

ENDLOCAL

PAUSE

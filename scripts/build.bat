@ECHO OFF
SETLOCAL

:: Build the pnfl-pdbtoexcel package.
:: Wipes the build/ staging dir first to avoid stale orphan files
:: from any source-file deletes or renames since the last build.

cd /d "%~dp0\.."

IF EXIST build (
    ECHO Removing stale build/...
    rmdir /s /q build
)

python -m build

ENDLOCAL

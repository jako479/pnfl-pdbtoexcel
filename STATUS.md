# pnfl-pdbtoexcel — Status

**Status: Complete**

Exports the contents of a WinLogStats database (.pdb) into a macro-enabled Excel worksheet (.xlsm).

## Implemented

- CLI (`pnfl convert-pdb`) with positional PDB/output args, optional offensive and defensive `.pln` plans (including second plans), config and play-path overrides, and skip-calcs/skip-totals flags
- WinLogStats `.pdb` binary parsing via ctypes, with validation of malformed data
- Joins per-team per-play stats to play-pool records for category metadata, using `fbpro98-gameplan` and `pnfl-playpool`
- Totals aggregation across teams, category filtering, and row sorting
- Excel workbook generation via XlsxWriter, including play worksheets, tendencies, and optional category worksheets
- `.xlsm` output with embedded prebuilt VBA macro project
- Layered config lookup (`convert-pdb.ini` / `.dev.ini` across CWD, project root, and package)

## Remaining

- Nothing outstanding for the current scope.

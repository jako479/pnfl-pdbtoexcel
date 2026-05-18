# pnfl-pdbtoexcel — Architecture

CLI tool that converts a WinLogStats `.pdb` (and optional FbPro '98 game plans) into a macro-enabled Excel workbook.

## Module layout

```
src/pnfl_pdbtoexcel/
├── __init__.py             # docstring
├── cli.py                  # argparse + main()
├── main.py                 # convert_pdb() orchestration
├── config.py               # Config dataclass, CategoryOrder, load_config(), load_category_order()
├── pdb.py                  # PDB binary format (ctypes) and parser
├── excel_workbook.py       # ExcelPdbWorkbook — low-level XlsxWriter row/sheet helpers
├── workbook_creator.py     # PdbWorkbookCreator — joins PDB stats to play pool, sorts, totals
└── resources/              # vbaProject*.bin — XLSM macro blocks
```

## What this package does

- Provides a CLI: `pnfl convert-pdb PDB OUTPUT [-o OFFENSE.pln] [-d DEFENSE.pln] [-o2 ...] [-d2 ...] [--play-path DIR] [--config FILE] [--skip-calcs] [--skip-totals]`
- Loads a WinLogStats `.pdb` via the `PDB` ctypes wrapper
- Optionally loads up to two offensive and two defensive `.pln` game plans (`fbpro98-gameplan`)
- Joins per-team per-play stats to play-pool records (`pnfl-playpool`) for category metadata
- Aggregates totals across teams, applies category filters, and sorts rows
- Emits an `.xlsx` or `.xlsm` workbook via XlsxWriter; `.xlsm` outputs embed the macro block from `resources/`

## What this package assumes

- The PDB is a well-formed FbPro '98 WinLogStats database (ctypes layout matches the C struct)
- The play pool root contains valid `.ply` files; failure to scan raises errors from `pnfl-playpool`
- The optional `.pln` files are well-formed; failure to parse raises `InvalidGamePlanError` from `fbpro98-gameplan`

## What this package enforces

CLI-level (raise SystemExit via argparse):
- `pdbfile` exists and has `.pdb` extension
- `outputfile` has `.xlsx` or `.xlsm` extension
- Each `--plnfile-*` argument exists and has `.pln` extension
- `--config` exists and has `.ini` extension

Config (raise `ValueError` / propagate from `configparser`):
- Required `[Settings]` keys present
- `PlayPath` resolves to a real directory at `read_play_pool()` time

## What this package does NOT do

- Parse `.ply` play files (delegates transitively via `pnfl-playpool` → `fbpro98-play`)
- Parse `.pln` game plans (delegates to `fbpro98-gameplan`)
- Emit non-Excel formats — XlsxWriter is the only output backend
- Produce a play catalog (lives in `pnfl-playcatalog`)

## Output contract

- `.xlsx` — plain XlsxWriter output
- `.xlsm` — same content with the VBA project from `resources/vbaProject*.bin` injected to provide PNFL's workbook macros

The two `.bin` resources are pre-built and shipped with the package; rebuilding them is out of scope.

## Testing

- `tests/test_cli.py` — argparse contract and `main()` exit codes
- `tests/test_pdb_parsing.py` — PDB binary parsing on real fixtures in `tests/data/`
- `tests/test_workbook_creation.py` — end-to-end workbook generation, asserted against expected XLSX outputs

PDB and PLN fixtures are real game-produced files; they are the authoritative ground truth for any wire-format question.

# pnfl-pdbtoexcel

Exports the contents of a WinLogStats database (.pdb) into an Excel worksheet (.xlsm).
Uses `fbpro98-gameplan` for `.pln` parsing and `pnfl-playpool` for play classification.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -e ..\fbpro98-play
pip install -e ..\fbpro98-gameplan
pip install -e ..\pnfl-playpool
pip install -e ".[dev]"
```

## Usage

```bash
python -m pnfl_pdbtoexcel PDB.pdb output.xlsm -d defense.pln -o offense.pln
```

Optional overrides:

```bash
python -m pnfl_pdbtoexcel PDB.pdb output.xlsm --config config/pdb_to_excel.ini
python -m pnfl_pdbtoexcel PDB.pdb output.xlsm --team Denver --pnfl-path E:\SIERRA\FbPro98\PNFL
```

Config lookup order:

1. `pdb_to_excel.ini` in the current working directory
2. `src/pnfl_pdbtoexcel/pdb_to_excel.ini`
3. `config/pdb_to_excel.ini`

## VBA Macros

When writing `.xlsm` output, a prebuilt VBA project is embedded:

- `vbaProject.bin`
- `vbaProject_categories.bin` (used when category worksheets are included)

See `excel-template/README.md` for how to update these.

## Testing

```bash
pytest
```

## Building a Release

Build the release zip:

```bash
py -3.13 scripts/build_release.py
```

This creates `dist/PdbToExcel-vX.Y.Z.zip`.

Smoke test in a clean venv:

1. Extract `dist/PdbToExcel-vX.Y.Z.zip` to a temp folder
2. Open a command prompt in the extracted folder

```bash
py -3.13 -m venv .venv
.venv\Scripts\activate
install.bat
python -m pnfl_pdbtoexcel "C:\path\to\test.pdb" "C:\path\to\test.xlsm"
deactivate
```

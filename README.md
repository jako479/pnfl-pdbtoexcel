# pnfl-pdbtoexcel

Exports the contents of a WinLogStats database (.pdb) into an Excel worksheet (.xlsm). Uses `fbpro98-gameplan` for `.pln` parsing and `pnfl-playpool` for play classification.

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

Distributed via the [`pnfl`](../pnfl) umbrella CLI:

```bash
pnfl convert-pdb PDB.pdb output.xlsm -d defense.pln -o offense.pln
```

Include a second defensive/offensive plan with `-d2`/`-o2`:

```bash
pnfl convert-pdb PDB.pdb output.xlsm -d defense.pln -d2 defense2.pln -o offense.pln -o2 offense2.pln
```

Optional overrides:

```bash
pnfl convert-pdb PDB.pdb output.xlsm --config config/convert-pdb.ini
pnfl convert-pdb PDB.pdb output.xlsm --play-path E:\SIERRA\FbPro98\PNFL
```

Config lookup order (first match wins; `.dev.ini` variants take precedence at each level):

1. `convert-pdb.dev.ini` / `convert-pdb.ini` in the current working directory
2. `config/convert-pdb.dev.ini` / `config/convert-pdb.ini` at the project root
3. `src/pnfl_pdbtoexcel/convert-pdb.dev.ini` / `src/pnfl_pdbtoexcel/convert-pdb.ini`

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

This project is distributed as part of the [`pnfl`](../pnfl) umbrella CLI. See `pnfl/scripts/build_release.py` for release packaging.

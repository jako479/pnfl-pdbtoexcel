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
1. `src/pnfl_pdbtoexcel/pdb_to_excel.ini`
2. `config/pdb_to_excel.ini`
3. `src/pnfl_pdbtoexcel/PdbToExcel.ini`
4. `config/PdbToExcel.ini`
## VBA Macros
When writing `.xlsm` output, a prebuilt VBA project is embedded:
- `vbaProject.bin`
- `vbaProject_categories.bin` (used when category worksheets are included)

These `.bin` files are extracted from a macro-enabled Excel workbook and embedded
with `xlsxwriter.Workbook.add_vba_project(...)`. If macros change, re-extract:
```bash
python .venv/Scripts/vba_extract.py path/to/source.xlsm
```
Store master `.xlsm` workbooks in `excel-template/`.
## Testing
```bash
pytest
```

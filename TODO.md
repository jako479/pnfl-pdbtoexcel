# TODO

## Immediate

No immediate TODOs.

## Next Steps

1. Extract the `.pln` reader/parser into `fbpro98-gameplan` as the shared library implementation.
2. Replace the local `PLN` parsing code in `pnfl-pdbtoexcel` with a thin adapter to that library.
3. Decide what the public game plan reader API should return:
   - raw records
   - domain objects
   - or both
4. Add a small set of real-file validation cases for `.pln` parsing once the shared reader is in place.

### Extraction Notes

- `pnfl-pdbtoexcel` is the first concrete consumer for the shared `.pln` library.
- The current project only needs the `.pln` behavior required by `pnfl-pdbtoexcel`, especially normal-play lookup/slot handling used during workbook generation.
- Special teams and stock special slot handling are not currently needed by `pnfl-pdbtoexcel`, so they should not drive the first library API.
- Prefer designing the shared `.pln` API from real caller needs first, then expanding it later for additional tools.

## Project Setup / Maintenance

1. Keep the master `.xlsm` workbooks to `excel-template/`:
   - `PdbToExcel.xlsm`
   - `PdbToExcelCategories.xlsm`
2. Regenerate `vbaProject.bin` and `vbaProject_categories.bin` from the master workbooks as needed.
3. Keep the VBA regeneration workflow documented in `docs/vba-workflow.md`.

## Nice To Have

1. Consider splitting `PdbToExcel.py` into smaller modules once the `.pln` parsing is moved out.
2. Tighten the README setup/usage notes after the GitHub move.
3. Add a thin `cli.py` entry once the different components have been extracted to shared modules and/or libraries.
4. Rename `PdbToExcel.py` to a more Pythonic module name/structure once the CLI and core components are separated cleanly.

# TODO

## Immediate

1. Get `pnfl-pdbtoexcel` into GitHub.
2. Sanity-check the current `PLN` offsets-table parsing against one or more real game plan files.
3. Confirm the slot handling for:
   - normal plays
   - special teams plays
   - stock special plays

## Next Steps

1. Extract the `.pln` reader/parser into `fbpro98-gameplan` as the shared library implementation.
2. Replace the local `PLN` parsing code in `pnfl-pdbtoexcel` with a thin adapter to that library.
3. Decide what the public game plan reader API should return:
   - raw records
   - domain objects
   - or both
4. Add a small set of real-file validation cases for `.pln` parsing once the shared reader is in place.

## Project Setup / Maintenance

1. Keep the master `.xlsm` workbooks to `excel-template/`:
   - `PdbToExcel.xlsm`
   - `PdbToExcelCategories.xlsm`
2. Regenerate `vbaProject.bin` and `vbaProject_categories.bin` from the master workbooks as needed.
3. Keep the VBA regeneration workflow documented in `docs/vba-workflow.md`.

## Nice To Have

1. Consider splitting `PdbToExcel.py` into smaller modules once the `.pln` parsing is moved out.
2. Tighten the README setup/usage notes after the GitHub move.

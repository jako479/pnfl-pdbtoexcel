# TODO

## Immediate

No immediate TODOs.

## Next Steps

1. Decide when `pnfl-pdbtoexcel` should declare `fbpro98-gameplan` as a package dependency instead of relying on sibling editable installs during development.
2. Add a small set of real `.pln` fixtures once redistributable sample files are available.
3. Decide whether `fbpro98-gameplan` should stay compatibility-shaped around `PLN`/`PlayInPlan` or expose a cleaner primary API with a thin adapter layer here.
4. Split more of `PdbToExcel.py` into focused modules now that `.pln` parsing is out of the file.

### Completed

- Extracted the `.pln` reader into `fbpro98-gameplan`.
- Replaced the local `PLN` parser in `pnfl-pdbtoexcel` with library imports.
- Added focused parser tests in `fbpro98-gameplan`.
- Kept the initial shared API shaped around real `pnfl-pdbtoexcel` needs: slot lookup and normal-play lookup by name.

## Project Setup / Maintenance

1. Keep the master `.xlsm` workbooks to `excel-template/`:
   - `PdbToExcel.xlsm`
   - `PdbToExcelCategories.xlsm`
2. Regenerate `vbaProject.bin` and `vbaProject_categories.bin` from the master workbooks as needed.
3. Keep the VBA regeneration workflow documented in `docs/vba-workflow.md`.

## Nice To Have

1. Rename `PdbToExcel.py` to a more Pythonic module name once the remaining responsibilities are split out.
2. Add a thin `cli.py` entry once the CLI and workbook-generation code are separated cleanly.
3. Expand test coverage around workbook generation paths that use offensive and defensive game plans together.

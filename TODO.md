# TODO

## Immediate

1. Split more of `pdb_to_excel.py` into focused modules.
2. Add a small set of real `.pln` fixtures once redistributable sample files are available.

## Long Term

1. Decide when `pnfl-pdbtoexcel` should declare `fbpro98-gameplan` as a package dependency instead of relying on sibling editable installs during development.
2. Decide whether `fbpro98-gameplan` should stay compatibility-shaped around `PLN`/`PlayInPlan` or expose a cleaner primary API with a thin adapter layer here.
3. Add a thin `cli.py` entry if the CLI and workbook-generation code are separated cleanly.
4. Expand test coverage around workbook generation paths that use offensive and defensive game plans together.

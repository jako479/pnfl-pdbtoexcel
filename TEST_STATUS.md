# pnfl-pdbtoexcel — Test Status

**Test Status: Tests Complete**

## Covered by automated tests

- CLI argument parsing: required positionals, all options, and extension validation
- Config and category-order loading, including play-path overrides and default fallbacks
- PDB binary parsing against a real game-produced fixture, with a plays snapshot and invalid-data rejection
- Category-order validation (missing categories and missing play types)
- End-to-end workbook generation with and without gameplans, asserting sheet names and row counts
- Skip-calcs/skip-totals behavior
- Cell-value verification across every column of the play, tendency, and category worksheets

## Needs tests

- Nothing outstanding for the current scope.

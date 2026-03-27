# VBA Workflow

`pnfl-pdbtoexcel` does not build VBA code itself. When writing `.xlsm` output,
it uses `xlsxwriter.Workbook.add_vba_project(...)` to embed a prebuilt VBA
project binary into the generated workbook.

## Files Used By This Project

- `src/pnfl_pdbtoexcel/vbaProject.bin`
- `src/pnfl_pdbtoexcel/vbaProject_categories.bin`

`PdbToExcel.py` chooses between them based on whether category-stats output is
enabled.

More specifically:

- `vbaProject.bin` is used for the normal spreadsheet output.
- `vbaProject_categories.bin` is used for the spreadsheet option that includes
  the category worksheets.

## Source Of Truth

The source of truth for the VBA is a macro-enabled Excel workbook (`.xlsm`) that
you maintain in Excel. Store those master workbooks in `excel-template/`.

Typical pattern:

- one source workbook for the normal macro project
- one source workbook for the category-stats macro project, if that VBA differs

If both output modes use the same VBA project, only one source workbook is
needed.

## How To Regenerate A VBA Binary

After editing the VBA in Excel and saving the source workbook as `.xlsm`, extract
 the VBA project with `xlsxwriter`'s helper script:

```powershell
.\.venv\Scripts\python.exe .\.venv\Scripts\vba_extract.py path\to\source.xlsm
```

That writes `vbaProject.bin` to the current working directory.

If the workbook contains a signed VBA project, the helper may also extract:

```text
vbaProjectSignature.bin
```

This project does not currently use the signature file.

## Updating This Project

1. Open the source `.xlsm` in Excel.
2. Update the VBA macros.
3. Save the workbook.
4. Run `vba_extract.py` against that workbook.
5. Copy the extracted `vbaProject.bin` into:
   `src/pnfl_pdbtoexcel/vbaProject.bin`

For the category-stats variant, repeat the same process and copy the extracted
binary into:

- `src/pnfl_pdbtoexcel/vbaProject_categories.bin`

## Recommended Working Convention

Keep the source `.xlsm` workbooks in `excel-template/` and treat the extracted
`.bin` files in this project as build assets derived from those workbooks.

Default convention:

- `excel-template/PdbToExcel.xlsm`
- `excel-template/PdbToExcelCategories.xlsm`

Then regenerate the binaries with:

```powershell
.\.venv\Scripts\python.exe .\.venv\Scripts\vba_extract.py .\excel-template\PdbToExcel.xlsm
Move-Item .\vbaProject.bin .\src\pnfl_pdbtoexcel\vbaProject.bin -Force
```

```powershell
.\.venv\Scripts\python.exe .\.venv\Scripts\vba_extract.py .\excel-template\PdbToExcelCategories.xlsm
Move-Item .\vbaProject.bin .\src\pnfl_pdbtoexcel\vbaProject_categories.bin -Force
```

## Important Notes

- `xlsxwriter` embeds the VBA project; it does not create or edit the VBA code.
- The `.bin` files should be replaced only when the source VBA changes.
- If generated `.xlsm` files stop exposing the expected macro behavior, verify
  that the correct extracted `.bin` was copied into `src/pnfl_pdbtoexcel`.

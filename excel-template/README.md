# Excel Templates

Store the master macro-enabled Excel workbooks for `pnfl-pdbtoexcel` here.

Recommended files:

- `PdbToExcel.xlsm`
- `PdbToExcelCategories.xlsm`

Purpose:

- `PdbToExcel.xlsm` is the source workbook for `src/pnfl_pdbtoexcel/vbaProject.bin`
- `PdbToExcelCategories.xlsm` is the source workbook for
  `src/pnfl_pdbtoexcel/vbaProject_categories.bin`

When the VBA changes:

1. Open the relevant `.xlsm` workbook in Excel.
2. Update the VBA.
3. Save the workbook.
4. Extract a fresh `vbaProject.bin` with:

```powershell
.\.venv\Scripts\python.exe .\.venv\Scripts\vba_extract.py .\excel-template\PdbToExcel.xlsm
Move-Item .\vbaProject.bin .\src\pnfl_pdbtoexcel\vbaProject.bin -Force
```

For the category-worksheet version:

```powershell
.\.venv\Scripts\python.exe .\.venv\Scripts\vba_extract.py .\excel-template\PdbToExcelCategories.xlsm
Move-Item .\vbaProject.bin .\src\pnfl_pdbtoexcel\vbaProject_categories.bin -Force
```

See also:

- `docs/vba-workflow.md`

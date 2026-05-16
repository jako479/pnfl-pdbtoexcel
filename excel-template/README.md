# VBA Templates

Master macro-enabled workbooks for pnfl-pdbtoexcel. The `.bin` files in `src/pnfl_pdbtoexcel/` are extracted from these and embedded into generated `.xlsm` output by xlsxwriter.

| Template                        | Binary                      | Used when              |
| ------------------------------- | --------------------------- | ---------------------- |
| `PdbToExcel.xlsm`               | `vbaProject.bin`            | Normal output          |
| `PdbToExcelWithCategories.xlsm` | `vbaProject_categories.bin` | Category stats enabled |

## Updating VBA

1. Edit VBA in the `.xlsm`, save.
2. Extract and replace the binary:

```powershell
..\.venv\Scripts\python.exe ..\.venv\Scripts\vba_extract.py .\excel-template\PdbToExcel.xlsm
Move-Item .\vbaProject.bin .\src\pnfl_pdbtoexcel\resources\vbaProject.bin -Force
```

For categories:

```powershell
..\.venv\Scripts\python.exe ..\.venv\Scripts\vba_extract.py .\excel-template\PdbToExcelWithCategories.xlsm
Move-Item .\vbaProject.bin .\src\pnfl_pdbtoexcel\resources\vbaProject_categories.bin -Force
```

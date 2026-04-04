PdbToExcel
==========

Converts WinLogStats game statistics (.pdb) into Excel spreadsheets (.xlsm).


REQUIREMENTS
------------

Python 3.10 or later.
Download from: https://www.python.org/downloads/

IMPORTANT: During Python installation, check the box that says
"Add Python to PATH". Without this, the install and run scripts
will not work.


INSTALLATION
------------

1. Extract this zip to a folder on your computer.
2. Double-click install.bat and wait for it to finish.
3. Open pdb_to_excel.ini in Notepad and set your team name and PNFL path.

You only need to run install.bat once.


USAGE
-----

Double-click PdbToExcel.bat, or run from the command line:

    PdbToExcel.bat "C:\path\to\stats.pdb" "C:\path\to\output.xlsm"

With game plans:

    PdbToExcel.bat "C:\path\to\stats.pdb" "C:\path\to\output.xlsm" -d "C:\path\to\defense.pln" -o "C:\path\to\offense.pln"

The output file extension matters: use .xlsm for macro-enabled sorting.


TROUBLESHOOTING
---------------

"python is not recognized" or "pip is not recognized":
    Python was installed without the PATH option. Reinstall Python and
    check "Add Python to PATH".

"No module named pnfl_pdbtoexcel":
    Run install.bat again.

install.bat shows errors:
    Make sure you have an internet connection (xlsxwriter is downloaded
    from the internet during install).

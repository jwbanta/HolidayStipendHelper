# HolidayStipendHelper

A Python command-line application being developed to help confirm which school
employees appear in one or more other lists. Matching and confidence assessment
will be designed after the source files have been inspected.

## File discovery

The discovery command inventories Excel workbooks (`.xlsx` and `.xlsm`) and
UTF-8 CSV files, including CSV exports from Google Sheets. It reports each
worksheet's name and visibility, row count, and maximum column count. CSV files
are treated as single tables. Rows and columns include blank cells; Excel files
may include extra blank rows or columns due to worksheet formatting.

The tool does not assume that the first row contains headers. To inspect
potential headers and data, request a limited number of sample rows explicitly.
Samples can contain personal information, so they are not shown by default.
Formulas are displayed as formulas, not calculated values. The files are read
without modification.

### Setup and use

Requires Python 3.10 or later.

```sh
python -m venv .venv
source .venv/bin/activate
python -m pip install -e .

holiday-stipend discover employees.xlsx comparison.csv
holiday-stipend discover --sample-rows 3 employees.xlsx comparison.csv
```

The command reports errors for individual files and exits unsuccessfully if
any file could not be inspected. It does not support legacy `.xls` files or
direct Google Sheets access yet; export a Google Sheet as CSV first.

## Tests

```sh
python -m unittest discover -s tests -v
```

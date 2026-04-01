# Employee Update Tool

## Folder Setup

Use this folder:

```text
employee update py
```

The folder should contain:

- `compare_employees.py`
- `preview_reports.py`
- `requirements.txt`
- `system_clean.xlsx` or `system_clean.xls`
- `hr_clean.xlsx` or `hr_clean.xls`

Note:

- if you use `py compare_employees.py`, the input files should be inside the `employee update py` folder
- if you use `py preview_reports.py` and upload files in the browser, the files do not need to be moved into the folder

## Input File Rules

### System File

Accepted names:

- `system_clean.xlsx`
- `system_clean.xls`

The system file should have a valid employee ID column and a status column.

Examples:

- `Employee ID`
- `Status`

### HR File

Accepted names:

- `hr_clean.xlsx`
- `hr_clean.xls`

The HR file must have the real header on row 5.

Expected HR columns include:

- `IDNO`
- `LASTNAME`
- `FIRSTNAME`
- `MIDDLE NAME`
- `ADDRESS`
- `JOBSTATUS`
- `BIRTHDAY`

Example HR header row:

```text
PROJECT | DEPTCODE | SITE | IDNO | LASTNAME | FIRSTNAME | MIDDLE NAME | ADDRESS | POSITION | POSITION LEVEL | JOBSTATUS | STATUS DATE | ... | BIRTHDAY | ...
```

## Install Libraries

Install the needed Python packages:

```bash
py -m pip install -r requirements.txt
```

## Run The Comparison

Generate the Excel reports:

```bash
py compare_employees.py
```

This creates:

- `inactive_to_update.xlsx`
- `active_to_update.xlsx`
- `new_active_employees.xlsx`

If one of those Excel files is open, the script saves to a fallback name like `_1.xlsx`.

## Run The Local Preview

Start the local web preview:

```bash
py preview_reports.py
```

Then open this in your browser:

```text
http://127.0.0.1:8000
```

From the preview page, you can now:

- upload a new `system` file
- upload a new `hr` file
- run the comparison again in the browser
- preview the new reports immediately
- use files from any local folder without copying them into `employee update py`

## What The Reports Mean

- `inactive_to_update.xlsx`: employee exists in both files, system says `active`, HR says not active
- `active_to_update.xlsx`: employee exists in both files, system says not active, HR says `active`
- `new_active_employees.xlsx`: employee is `active` in HR but does not exist in system

## Quick Steps

1. Put `system_clean` and `hr_clean` Excel files inside the `employee update py` folder.
2. Make sure the HR file header is on row 5.
3. Install libraries with `py -m pip install -r requirements.txt`.
4. Run `py compare_employees.py`.
5. Run `py preview_reports.py`.
6. Open `http://127.0.0.1:8000` in your browser.
7. Optional: upload new Excel files in the browser to compare again without replacing the original files in the main folder.

## Dos And Donts

### Do

- keep the HR header on row 5
- keep employee IDs clean and consistent between files
- use `.xlsx` when possible
- close report files in Excel if you want the exact same output filename
- use the web upload if you want to test different files without overwriting your main inputs

### Don't

- do not rename important columns like `IDNO`, `JOBSTATUS`, `Employee ID`, or `Status` unless you also update the script
- do not add blank rows before the HR header row
- do not keep output Excel files open if you expect direct overwrite
- do not mix different employee ID formats for the same employee

# Dos And Donts

## Purpose

This guide explains the correct way to prepare files for the employee update tool and what to avoid.

## Do

### 1. Put the correct files in the same folder

Keep these files inside the `employee update py` folder:

- `compare_employees.py`
- `preview_reports.py`
- `requirements.txt`
- `system_clean.xlsx` or `system_clean.xls`
- `hr_clean.xlsx` or `hr_clean.xls`

Example:

```text
employee update py/
  compare_employees.py
  preview_reports.py
  requirements.txt
  system_clean.xlsx
  hr_clean.xls
```

Important:

- this folder setup is needed when using `py compare_employees.py`
- if using the browser upload in `py preview_reports.py`, users do not need to move the Excel files into the `employee update py` folder

### 2. Keep the HR header on row 5

The HR file must use row 5 as the real header row.

Example:

```text
Row 1: QM BUILDERS
Row 2: Employee Profile
Row 3: As of March 2026
Row 4: blank or title row
Row 5: PROJECT | DEPTCODE | SITE | IDNO | LASTNAME | FIRSTNAME | MIDDLE NAME | ADDRESS | POSITION | JOBSTATUS | BIRTHDAY
```

### 3. Keep employee IDs consistent

Employee IDs should match between the system file and HR file.

Good example:

```text
System: 001554
HR:     001554
```

### 4. Use supported column names

The tool expects these kinds of columns:

System file examples:

- `Employee ID`
- `Status`

HR file examples:

- `IDNO`
- `LASTNAME`
- `FIRSTNAME`
- `MIDDLE NAME`
- `ADDRESS`
- `JOBSTATUS`
- `BIRTHDAY`

### 5. Prefer `.xlsx` for new files

`.xlsx` is easier to work with than old `.xls` files.

Good example:

```text
system_clean.xlsx
hr_clean.xlsx
```

### 6. Close Excel output files before rerunning

If an output file is open in Excel, Windows may block overwrite.

Example:

```text
inactive_to_update.xlsx is open in Excel
```

Result:

```text
The script may save as inactive_to_update_1.xlsx instead
```

### 7. Use the browser upload for testing other files

If you want to compare a different pair of files without replacing the originals, use the upload form in the local preview page.

Example:

```text
http://127.0.0.1:8000
```

Upload:

- one system file
- one HR file

Then click:

```text
Upload And Compare
```

Example:

```text
Your files can stay in Downloads, Desktop, or any local folder.
You just select them in the browser upload form.
```

## Don't

### 1. Do not move the HR header to another row

Bad example:

```text
Row 1: PROJECT | DEPTCODE | SITE | IDNO | LASTNAME | FIRSTNAME | JOBSTATUS
```

Why:

The current tool expects the HR header on row 5.

### 2. Do not rename important columns without updating the script

Bad example:

```text
IDNO -> EMPLOYEE NUMBER
JOBSTATUS -> CURRENT STATE
```

Why:

The script may not recognize those renamed columns.

### 3. Do not use different employee ID formats for the same employee

Bad example:

```text
System: 1554
HR:     001554
```

Why:

They may be treated as different employees.

### 4. Do not leave report files open if you expect overwrite

Bad example:

```text
new_active_employees.xlsx is open while running compare_employees.py
```

Why:

You may get:

```text
new_active_employees_1.xlsx
```

instead of replacing the original file.

### 5. Do not upload the wrong file type

Bad example:

- PDF
- CSV with different columns
- image file

Why:

The tool is designed for Excel `.xls` and `.xlsx` files.

### 6. Do not insert many extra blank rows before the HR data

Bad example:

```text
Rows 1-10: blank
Row 11: actual header
```

Why:

The tool is currently configured for HR header row 5.

## Recommended Workflow

1. Prepare `system_clean` and `hr_clean` Excel files.
2. Make sure the HR header is on row 5.
3. Install libraries:

```bash
py -m pip install -r requirements.txt
```

4. Run the compare script:

```bash
py compare_employees.py
```

5. Run the preview site:

```bash
py preview_reports.py
```

6. Open:

```text
http://127.0.0.1:8000
```

7. Use the download buttons to save the generated reports.

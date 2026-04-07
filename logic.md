# Logic

## Compare Employees

The main comparison flow uses `system_clean` and `hr_clean` to produce:

- `inactive_to_update.xlsx`
- `active_to_update.xlsx`
- `new_active_employees.xlsx`

### How `system_clean` is read

- The app does **not** rely on fixed Excel column letters.
- It scans the first rows to find the correct header row.
- It detects columns by normalized header names and aliases.
- Important logical fields:
  - `id`
  - `system_status`
  - `department`

This means column order can change and the comparison can still work, as long as the headers are still recognizable.

### How `hr_clean` is read

- The app reads `hr_clean` using the expected header row configured in code.
- After that, it also normalizes header names.
- Important logical fields:
  - `id`
  - `hr_status`

This means column order can change, but if the HR file layout changes too much, especially the header row position, it may need adjustment.

### Comparison rules

- `Inactive To Update`
  - Employee exists in both files
  - `system_status` is `active`
  - `hr_status` is not `active`

- `Active To Update`
  - Employee exists in both files
  - `system_status` is not `active`
  - `hr_status` is `active`

- `New Active Employees`
  - Employee exists in `hr_clean`
  - Employee does not exist in `system_clean`
  - `hr_status` is `active`

### Why this is relatively safe

- The main comparison logic is based on detected headers and logical field names.
- It is not hardcoded to specific Excel column letters like `A`, `D`, or `H`.
- It is more resilient to column movement than a raw column-index approach.


## Updated `system_clean`

This is a separate flow from the main comparison flow.

It is used to calculate employee totals for the printable report second section:

- `Total Inactive Employees`
- `Total Active Employees`
- `Total Employees Overall`

### File structure expected

- Headers are read from Excel row `1`
- Data starts from Excel row `2`

### Column logic

This flow supports three levels of column selection:

1. Manual override from the dashboard dropdowns
2. Auto-detect by header name
3. Fallback defaults

Fallback defaults are:

- Status column: Excel column `D`
- Branch basis column: Excel column `H`

### Status logic

- Employee status is read from the selected status column.
- `active` counts as active.
- Any non-`active` value is counted as inactive for the status summary.

### Branch mapping logic

The uploaded updated `system_clean` uses a branch basis column, and that value is mapped into printable parent branches.

Mapped parent branches:

- `QM BUILDERS`
  - `qm realty`
  - `qmb production`
  - `qmb hardware`
  - `qmb equipment`
  - `qmb construction`
  - `qmb constructions`

- `ADAMANT`
  - `adc construction`
  - `adc constructions`

- `QM FARMS`
  - `qmb farm`
  - `qmb farms`

- `QM DIVING RESORT`
  - `cafe de casilda`
  - `diving resort`

- `QGDC`
  - `qgdc`
  - `qgdc construction`
  - `qgdc constructions`

- `QMAZ HOLDINGS`
  - `qmaz operations`

### Matching behavior

- Matching is normalized
- It supports both spaced and compact values

Examples:

- `qmb constructions`
- `qmbconstruction`

Both can map to the same printable parent branch.

### Unmapped rows

- If a branch basis value does not match the configured aliases, it is counted under:
  - `UNMAPPED / OTHER`

This prevents totals from disappearing silently.

### Printable report usage

The printable report has two sections:

- Section 1
  - employee update totals from the comparison reports

- Section 2
  - updated `system_clean` branch breakdown with:
    - inactive employees
    - active employees
    - overall employees

When `Add Status Totals To Print Report` is clicked, the branch breakdown from the updated `system_clean` flow is stored and shown in Section 2 of the printable report.

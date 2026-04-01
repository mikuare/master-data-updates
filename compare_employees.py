from pathlib import Path
import importlib
import re

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
HR_HEADER_ROW = 4
ID_ALIASES = {"employeeid", "idno", "employeeno"}
HR_STATUS_ALIASES = {"jobstatus", "hrstatus", "employmentstatus"}
SYSTEM_STATUS_ALIASES = {"status", "systemstatus"}
DISPLAY_COLUMN_RENAMES = {
    "lastname": "last_name",
    "firstname": "first_name",
    "middlename": "middle_name",
    "middlenameinitial": "middle_name",
    "birthdate": "birth_date",
    "birthday": "birth_date",
    "deptcode": "dept_code",
}


def find_input_file(*names: str) -> Path:
    for name in names:
        path = BASE_DIR / name
        if path.exists():
            return path
    tried = ", ".join(names)
    raise FileNotFoundError(
        f"Could not find any of these files in {BASE_DIR}: {tried}"
    )


def read_excel_file(path: Path, **kwargs) -> pd.DataFrame:
    if path.suffix.lower() == ".xls":
        try:
            importlib.import_module("xlrd")
        except ImportError as exc:
            raise ImportError(
                "The file "
                f"'{path.name}' is an old .xls workbook and requires the 'xlrd' package.\n"
                "Install it with:\n"
                "  py -m pip install xlrd>=2.0.1"
            ) from exc
        return pd.read_excel(path, engine="xlrd", **kwargs)

    return pd.read_excel(path, **kwargs)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        re.sub(r"[^a-z0-9]+", "", str(column).strip().lower())
        for column in df.columns
    ]
    return df


def normalize_label(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def find_header_row(
    path: Path,
    first_group: set[str],
    second_group: set[str],
    max_rows: int = 15,
) -> int:
    preview = read_excel_file(path, header=None)
    limit = min(len(preview.index), max_rows)

    for row_idx in range(limit):
        row_values = {
            normalize_label(value)
            for value in preview.iloc[row_idx].tolist()
            if normalize_label(value)
        }
        if row_values.intersection(first_group) and row_values.intersection(second_group):
            return row_idx

    return 0


def build_full_name(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for column in ["last_name", "first_name", "middle_name"]:
        if column not in df.columns:
            df[column] = ""
        df[column] = df[column].fillna("").astype(str).str.strip()

    df["full_name"] = (
        df["last_name"] + ", " + df["first_name"] + " " + df["middle_name"]
    ).str.replace(r"\s+", " ", regex=True).str.strip(" ,")
    return df


def select_report_columns(df: pd.DataFrame) -> pd.DataFrame:
    preferred = [
        "id",
        "full_name",
        "last_name",
        "first_name",
        "middle_name",
        "address",
        "birth_date",
        "system_status",
        "hr_status",
    ]
    existing = [column for column in preferred if column in df.columns]
    return df[existing]


def save_report(df: pd.DataFrame, filename: str) -> Path:
    target = BASE_DIR / filename

    try:
        df.to_excel(target, index=False)
        return target
    except PermissionError:
        for counter in range(1, 100):
            fallback = BASE_DIR / f"{target.stem}_{counter}{target.suffix}"
            try:
                df.to_excel(fallback, index=False)
                print(
                    f"Could not overwrite {target.name} because it is open. "
                    f"Saved to {fallback.name} instead."
                )
                return fallback
            except PermissionError:
                continue
        raise


print("Loading files...")

system_path = find_input_file("system_clean.xlsx", "system_clean.xls")
hr_path = find_input_file("hr_clean.xlsx", "hr_clean.xls")

system_header_row = find_header_row(system_path, ID_ALIASES, SYSTEM_STATUS_ALIASES)
hr_header_row = HR_HEADER_ROW

system = read_excel_file(system_path, header=system_header_row)
hr = read_excel_file(hr_path, header=hr_header_row)

# Normalize column names
system = normalize_columns(system)
hr = normalize_columns(hr)

system = system.rename(columns=DISPLAY_COLUMN_RENAMES)
hr = hr.rename(columns=DISPLAY_COLUMN_RENAMES)

# Rename columns to standard names
system = system.rename(columns={
    "employeeid": "id",
    "idno": "id",
    "status": "system_status",
})

hr = hr.rename(columns={
    "employeeid": "id",
    "idno": "id",
    "jobstatus": "hr_status",
})

required_system_columns = {"id", "system_status"}
required_hr_columns = {"id", "hr_status"}

missing_system = required_system_columns - set(system.columns)
missing_hr = required_hr_columns - set(hr.columns)

if missing_system:
    raise KeyError(
        f"Missing required columns in {system_path.name}: {sorted(missing_system)}. "
        f"Available columns: {list(system.columns)}"
    )

if missing_hr:
    raise KeyError(
        f"Missing required columns in {hr_path.name}: {sorted(missing_hr)}. "
        f"Available columns: {list(hr.columns)}"
    )

# Clean values
system["id"] = system["id"].astype(str).str.strip()
hr["id"] = hr["id"].astype(str).str.strip()
system["system_status"] = system["system_status"].astype(str).str.strip().str.lower()
hr["hr_status"] = hr["hr_status"].astype(str).str.strip().str.lower()

print("Comparing data...")

system_report = system[["id", "system_status"]].copy()
hr_report = hr.copy()
hr_report = build_full_name(hr_report)

matched_employees = system_report.merge(hr_report, on="id", how="inner")
new_active_employees = hr_report[
    (~hr_report["id"].isin(system_report["id"])) &
    (hr_report["hr_status"] == "active")
].copy()

# System is active, but HR says not active.
inactive_to_update = matched_employees[
    (matched_employees["system_status"] == "active") &
    (matched_employees["hr_status"] != "active")
].copy()

# System is inactive, but HR says active.
active_to_update = matched_employees[
    (matched_employees["system_status"] != "active") &
    (matched_employees["hr_status"] == "active")
].copy()

inactive_to_update = select_report_columns(inactive_to_update)
active_to_update = select_report_columns(active_to_update)
new_active_employees = select_report_columns(new_active_employees)

# Save results
inactive_output = save_report(inactive_to_update, "inactive_to_update.xlsx")
active_output = save_report(active_to_update, "active_to_update.xlsx")
new_active_output = save_report(new_active_employees, "new_active_employees.xlsx")

print("Done!")
print(f"System file: {system_path.name}")
print(f"HR file: {hr_path.name}")
print(f"Inactive report: {inactive_output.name}")
print(f"Active report: {active_output.name}")
print(f"New active report: {new_active_output.name}")
print(f"Matched employees checked: {len(matched_employees)}")
print(f"Set to inactive: {len(inactive_to_update)}")
print(f"Set to active: {len(active_to_update)}")
print(f"New active employees not in system: {len(new_active_employees)}")

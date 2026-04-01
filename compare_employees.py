from pathlib import Path
import importlib
import re

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
HR_HEADER_ROW = 4
ID_ALIASES = {"employeeid", "idno", "employeeno"}
SYSTEM_STATUS_ALIASES = {"status", "systemstatus"}
DISPLAY_COLUMN_RENAMES = {
    "lastname": "last_name",
    "firstname": "first_name",
    "middlename": "middle_name",
    "middlenameinitial": "middle_name",
    "birthdate": "birth_date",
    "birthday": "birth_date",
    "jobstatus": "hr_status",
    "deptcode": "dept_code",
}
REPORT_FILENAMES = {
    "inactive_to_update": "inactive_to_update.xlsx",
    "active_to_update": "active_to_update.xlsx",
    "new_active_employees": "new_active_employees.xlsx",
}


def find_input_file(*names: str, base_dir: Path = BASE_DIR) -> Path:
    for name in names:
        path = base_dir / name
        if path.exists():
            return path
    tried = ", ".join(names)
    raise FileNotFoundError(
        f"Could not find any of these files in {base_dir}: {tried}"
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


def save_report(df: pd.DataFrame, filename: str, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / filename

    try:
        df.to_excel(target, index=False)
        return target
    except PermissionError:
        for counter in range(1, 100):
            fallback = output_dir / f"{target.stem}_{counter}{target.suffix}"
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


def prepare_system_df(system_path: Path) -> pd.DataFrame:
    system_header_row = find_header_row(system_path, ID_ALIASES, SYSTEM_STATUS_ALIASES)
    system = read_excel_file(system_path, header=system_header_row)
    system = normalize_columns(system)
    system = system.rename(columns=DISPLAY_COLUMN_RENAMES)
    system = system.rename(columns={
        "employeeid": "id",
        "idno": "id",
        "status": "system_status",
    })

    required_columns = {"id", "system_status"}
    missing = required_columns - set(system.columns)
    if missing:
        raise KeyError(
            f"Missing required columns in {system_path.name}: {sorted(missing)}. "
            f"Available columns: {list(system.columns)}"
        )

    system["id"] = system["id"].astype(str).str.strip()
    system["system_status"] = (
        system["system_status"].astype(str).str.strip().str.lower()
    )
    return system


def prepare_hr_df(hr_path: Path) -> pd.DataFrame:
    hr = read_excel_file(hr_path, header=HR_HEADER_ROW)
    hr = normalize_columns(hr)
    hr = hr.rename(columns=DISPLAY_COLUMN_RENAMES)
    hr = hr.rename(columns={
        "employeeid": "id",
        "idno": "id",
    })

    required_columns = {"id", "hr_status"}
    missing = required_columns - set(hr.columns)
    if missing:
        raise KeyError(
            f"Missing required columns in {hr_path.name}: {sorted(missing)}. "
            f"Available columns: {list(hr.columns)}"
        )

    hr["id"] = hr["id"].astype(str).str.strip()
    hr["hr_status"] = hr["hr_status"].astype(str).str.strip().str.lower()
    return build_full_name(hr)


def generate_reports(system: pd.DataFrame, hr: pd.DataFrame) -> dict[str, pd.DataFrame]:
    system_report = system[["id", "system_status"]].copy()
    matched_employees = system_report.merge(hr, on="id", how="inner")

    inactive_to_update = matched_employees[
        (matched_employees["system_status"] == "active") &
        (matched_employees["hr_status"] != "active")
    ].copy()

    active_to_update = matched_employees[
        (matched_employees["system_status"] != "active") &
        (matched_employees["hr_status"] == "active")
    ].copy()

    new_active_employees = hr[
        (~hr["id"].isin(system_report["id"])) &
        (hr["hr_status"] == "active")
    ].copy()

    return {
        "inactive_to_update": select_report_columns(inactive_to_update),
        "active_to_update": select_report_columns(active_to_update),
        "new_active_employees": select_report_columns(new_active_employees),
        "_matched_count": pd.DataFrame([{"count": len(matched_employees)}]),
    }


def process_reports(
    system_path: Path | None = None,
    hr_path: Path | None = None,
    output_dir: Path = BASE_DIR,
) -> dict[str, object]:
    if system_path is None:
        system_path = find_input_file(
            "system_clean.xlsx",
            "system_clean.xls",
            base_dir=BASE_DIR,
        )
    if hr_path is None:
        hr_path = find_input_file(
            "hr_clean.xlsx",
            "hr_clean.xls",
            base_dir=BASE_DIR,
        )

    system = prepare_system_df(system_path)
    hr = prepare_hr_df(hr_path)
    reports = generate_reports(system, hr)

    inactive_output = save_report(
        reports["inactive_to_update"],
        REPORT_FILENAMES["inactive_to_update"],
        output_dir,
    )
    active_output = save_report(
        reports["active_to_update"],
        REPORT_FILENAMES["active_to_update"],
        output_dir,
    )
    new_active_output = save_report(
        reports["new_active_employees"],
        REPORT_FILENAMES["new_active_employees"],
        output_dir,
    )

    return {
        "system_file": system_path.name,
        "hr_file": hr_path.name,
        "inactive_output": inactive_output,
        "active_output": active_output,
        "new_active_output": new_active_output,
        "matched_count": int(reports["_matched_count"].iloc[0]["count"]),
        "inactive_count": len(reports["inactive_to_update"]),
        "active_count": len(reports["active_to_update"]),
        "new_active_count": len(reports["new_active_employees"]),
    }


def main() -> None:
    print("Loading files...")
    result = process_reports()
    print("Comparing data...")
    print("Done!")
    print(f"System file: {result['system_file']}")
    print(f"HR file: {result['hr_file']}")
    print(f"Inactive report: {Path(result['inactive_output']).name}")
    print(f"Active report: {Path(result['active_output']).name}")
    print(f"New active report: {Path(result['new_active_output']).name}")
    print(f"Matched employees checked: {result['matched_count']}")
    print(f"Set to inactive: {result['inactive_count']}")
    print(f"Set to active: {result['active_count']}")
    print(f"New active employees not in system: {result['new_active_count']}")


if __name__ == "__main__":
    main()

from pathlib import Path
from datetime import datetime
import importlib
import re
import json
import shutil
from zoneinfo import ZoneInfo

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
BRANCH_HISTORY_DIR = BASE_DIR / "branch_history"
HR_HEADER_ROW = 4
ID_ALIASES = {"employeeid", "idno", "employeeno"}
SYSTEM_STATUS_ALIASES = {"status", "systemstatus"}
SYSTEM_BRANCH_ALIASES = {
    "branch",
    "branchbasis",
    "branchgroup",
    "grouping",
    "parentbranch",
    "company",
    "division",
    "group",
}
DISPLAY_COLUMN_RENAMES = {
    "lastname": "last_name",
    "firstname": "first_name",
    "middlename": "middle_name",
    "middlenameinitial": "middle_name",
    "employeename": "full_name",
    "employeefullname": "full_name",
    "fullname": "full_name",
    "birthdate": "birth_date",
    "birthday": "birth_date",
    "jobstatus": "hr_status",
    "deptcode": "dept_code",
}
REPORT_FILENAMES = {
    "inactive_to_update": "inactive_to_update.xlsx",
    "active_to_update": "active_to_update.xlsx",
    "new_active_employees": "new_active_employees.xlsx",
    "new_in_system_since_last_month": "new_in_system_since_last_month.xlsx",
}
APP_TIMEZONE = ZoneInfo("Asia/Manila")


def local_now() -> datetime:
    return datetime.now(APP_TIMEZONE)


SYSTEM_BRANCH_GROUPS = {
    "QM BUILDERS": [
        "qm builders",
        "qm realty",
        "qmb production",
        "qmb hardware",
        "qmb equipment",
        "qmb construction",
        "qmb constructions",
    ],
    "ADAMANT": [
        "adamant",
        "adc construction",
        "adc constructions",
    ],
    "QM FARMS": [
        "qm farms",
        "qmb farm",
        "qmb farms",
    ],
    "QM DIVING RESORT": [
        "qm diving resort",
        "cafe de casilda",
        "diving resort",
    ],
    "QGDC": [
        "qgdc",
        "qgdc construction",
        "qgdc constructions",
    ],
    "QMAZ HOLDINGS": [
        "qmaz holdings",
        "qmaz operations",
    ],
}
UNMAPPED_BRANCH_LABEL = "UNMAPPED / OTHER"


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


def normalize_branch_name(value: object) -> str:
    return " ".join(str(value).strip().lower().split())


def compact_branch_name(value: object) -> str:
    return "".join(ch for ch in normalize_branch_name(value) if ch.isalnum())


def map_system_branch_label(value: object) -> str:
    normalized = normalize_branch_name(value)
    compact = compact_branch_name(value)
    if not compact:
        return UNMAPPED_BRANCH_LABEL

    for canonical, aliases in SYSTEM_BRANCH_GROUPS.items():
        for alias in aliases:
            normalized_alias = normalize_branch_name(alias)
            compact_alias = compact_branch_name(alias)
            if normalized_alias and normalized_alias in normalized:
                return canonical
            if compact_alias and compact_alias in compact:
                return canonical
            if compact and compact == compact_alias:
                return canonical

    return UNMAPPED_BRANCH_LABEL


def slugify_branch_key(value: str) -> str:
    normalized = re.sub(r"[^A-Z0-9]+", "_", value.strip().upper()).strip("_")
    return normalized or "UNMAPPED_OTHER"


def resolve_system_branch_column(system: pd.DataFrame) -> str | None:
    available_columns = list(system.columns)
    for column in available_columns:
        if normalize_label(column) in SYSTEM_BRANCH_ALIASES:
            return str(column)
    if len(available_columns) >= 8:
        return str(available_columns[7])
    return None


def determine_parent_branch_keys(system: pd.DataFrame) -> list[str]:
    if "parent_branch" not in system.columns:
        return [UNMAPPED_BRANCH_LABEL]
    branch_counts = (
        system["parent_branch"]
        .astype(str)
        .str.strip()
        .replace("", UNMAPPED_BRANCH_LABEL)
        .value_counts()
    )
    if branch_counts.empty:
        return [UNMAPPED_BRANCH_LABEL]
    ordered = [str(key) for key in branch_counts.index.tolist()]
    non_unmapped = [key for key in ordered if key != UNMAPPED_BRANCH_LABEL]
    if non_unmapped:
        return non_unmapped + ([UNMAPPED_BRANCH_LABEL] if UNMAPPED_BRANCH_LABEL in ordered else [])
    return ordered


def get_hr_branch_label(hr_path: Path) -> str:
    try:
        preview = read_excel_file(hr_path, header=None, nrows=1).fillna("")
    except Exception:
        return ""
    if preview.empty:
        return ""
    row_values = [str(value).strip() for value in preview.iloc[0].tolist()]
    non_empty = [value for value in row_values if value and value.lower() != "nan"]
    return " | ".join(non_empty)


def determine_run_branch_key(system: pd.DataFrame, hr_path: Path) -> str:
    hr_branch_label = get_hr_branch_label(hr_path)
    system_branches = [branch for branch in determine_parent_branch_keys(system) if branch != UNMAPPED_BRANCH_LABEL]
    mapped_hr_branch = map_system_branch_label(hr_branch_label)
    if mapped_hr_branch != UNMAPPED_BRANCH_LABEL:
        if system_branches and mapped_hr_branch not in system_branches:
            raise ValueError(
                "The uploaded hr_clean branch does not match the detected parent branch in system_clean. "
                f"HR branch: {mapped_hr_branch}. System branches found: {', '.join(system_branches)}."
            )
        return mapped_hr_branch

    if len(system_branches) == 1:
        return system_branches[0]
    if not hr_branch_label:
        raise ValueError(
            "Could not determine the run branch because hr_clean row 1 is empty or unreadable."
        )
    raise ValueError(
        "Could not confidently determine the run branch from hr_clean row 1. "
        f"HR row 1: {hr_branch_label}. "
        f"System branches found: {', '.join(system_branches) if system_branches else UNMAPPED_BRANCH_LABEL}."
    )


def get_branch_history_dir(branch_key: str) -> Path:
    return BRANCH_HISTORY_DIR / slugify_branch_key(branch_key)


def list_branch_snapshot_dirs(branch_key: str) -> list[Path]:
    history_dir = get_branch_history_dir(branch_key)
    if not history_dir.exists():
        return []
    return sorted(
        [path for path in history_dir.iterdir() if path.is_dir()],
        key=lambda path: path.name,
        reverse=True,
    )


def get_snapshot_period_key(value: datetime | None = None) -> str:
    return (value or local_now()).strftime("%Y%m")


def filter_snapshot_dirs_before_period(
    snapshot_dirs: list[Path],
    period_key: str,
) -> list[Path]:
    return [path for path in snapshot_dirs if path.name[:6] < period_key]


def load_snapshot_ids(snapshot_dir: Path) -> set[str]:
    ids_file = snapshot_dir / "system_ids.json"
    if not ids_file.exists():
        return set()
    try:
        data = json.loads(ids_file.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return set()
    if not isinstance(data, dict):
        return set()
    ids = data.get("ids", [])
    if not isinstance(ids, list):
        return set()
    return {str(item).strip() for item in ids if str(item).strip()}


def load_all_historical_ids_before_period(period_key: str) -> set[str]:
    if not BRANCH_HISTORY_DIR.exists():
        return set()
    all_ids: set[str] = set()
    for branch_dir in BRANCH_HISTORY_DIR.iterdir():
        if not branch_dir.is_dir():
            continue
        for snapshot_dir in branch_dir.iterdir():
            if not snapshot_dir.is_dir():
                continue
            if snapshot_dir.name[:6] >= period_key:
                continue
            all_ids.update(load_snapshot_ids(snapshot_dir))
    return all_ids


def save_branch_snapshot(
    branch_key: str,
    system_path: Path,
    system: pd.DataFrame,
    output_dir: Path,
    hr_path: Path,
) -> Path:
    period_key = get_snapshot_period_key()
    existing_for_period = [
        path
        for path in list_branch_snapshot_dirs(branch_key)
        if path.name[:6] == period_key
    ]
    if existing_for_period:
        return existing_for_period[0]

    timestamp = local_now().strftime("%Y%m%d_%H%M%S")
    snapshot_dir = get_branch_history_dir(branch_key) / timestamp
    snapshot_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(system_path, snapshot_dir / "system_clean_snapshot.xlsx")
    ids = sorted(
        {
            str(value).strip()
            for value in system["id"].astype(str).tolist()
            if str(value).strip()
        }
    )
    (snapshot_dir / "system_ids.json").write_text(
        json.dumps({"ids": ids}, indent=2),
        encoding="utf-8",
    )
    (snapshot_dir / "run_info.json").write_text(
        json.dumps(
            {
                "branch_key": branch_key,
                "saved_at": local_now().isoformat(),
                "system_file": system_path.name,
                "hr_file": hr_path.name,
                "output_dir": str(output_dir),
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    return snapshot_dir


def build_new_in_system_report_for_branch(
    system: pd.DataFrame,
    hr: pd.DataFrame,
    branch_key: str,
) -> tuple[pd.DataFrame, dict[str, object]]:
    branch_system = system[system["parent_branch"] == branch_key].copy()
    if branch_system.empty:
        empty = system.iloc[0:0].copy()
        empty["hr_status"] = ""
        return select_report_columns(empty), {
            "count": 0,
            "active_count": 0,
            "inactive_count": 0,
            "previous_snapshot_found": False,
            "previous_snapshot_name": "",
            "missing_snapshot_branches": [branch_key],
            "branches_with_snapshot": 0,
        }

    current_period_key = get_snapshot_period_key()
    all_historical_ids = load_all_historical_ids_before_period(current_period_key)
    previous_snapshots = filter_snapshot_dirs_before_period(
        list_branch_snapshot_dirs(branch_key),
        current_period_key,
    )
    if not previous_snapshots:
        empty = branch_system.iloc[0:0].copy()
        empty["hr_status"] = ""
        return select_report_columns(empty), {
            "count": 0,
            "active_count": 0,
            "inactive_count": 0,
            "previous_snapshot_found": False,
            "previous_snapshot_name": "",
            "missing_snapshot_branches": [branch_key],
            "branches_with_snapshot": 0,
        }

    previous_snapshot = previous_snapshots[0]
    previous_ids = load_snapshot_ids(previous_snapshot)
    branch_system["id"] = branch_system["id"].astype(str).str.strip()
    branch_system = branch_system[branch_system["id"] != ""].copy()
    new_in_system = branch_system[
        ~branch_system["id"].isin(previous_ids) &
        ~branch_system["id"].isin(all_historical_ids)
    ].copy()

    hr_subset = hr.copy()
    for column in ["id", "hr_status"]:
        if column not in hr_subset.columns:
            hr_subset[column] = ""
    if not new_in_system.empty:
        new_in_system = new_in_system.merge(
            hr_subset[["id", "hr_status"]],
            on="id",
            how="left",
        )
        new_in_system["hr_status"] = new_in_system["hr_status"].fillna("").astype(str).str.strip().str.lower()
    else:
        new_in_system["hr_status"] = ""

    active_count = int((new_in_system["system_status"] == "active").sum()) if "system_status" in new_in_system.columns else 0
    inactive_count = int(len(new_in_system.index) - active_count)
    return select_report_columns(new_in_system), {
        "count": len(new_in_system.index),
        "active_count": active_count,
        "inactive_count": inactive_count,
        "previous_snapshot_found": True,
        "previous_snapshot_name": f"{branch_key}: {previous_snapshot.name}",
        "missing_snapshot_branches": [],
        "branches_with_snapshot": 1,
    }


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
    if "full_name" in df.columns:
        df["full_name"] = df["full_name"].fillna("").astype(str).str.strip()
    else:
        df["full_name"] = ""
    for column in ["last_name", "first_name", "middle_name"]:
        if column not in df.columns:
            df[column] = ""
        df[column] = df[column].fillna("").astype(str).str.strip()

    derived_name = (
        df["last_name"] + ", " + df["first_name"] + " " + df["middle_name"]
    ).str.replace(r"\s+", " ", regex=True).str.strip(" ,")
    df["full_name"] = df["full_name"].where(df["full_name"] != "", derived_name)
    return df


def select_report_columns(df: pd.DataFrame) -> pd.DataFrame:
    if "department" not in df.columns:
        df = df.copy()
        df["department"] = ""

    preferred = [
        "id",
        "full_name",
        "last_name",
        "first_name",
        "middle_name",
        "dept_code",
        "department",
        "position",
        "address",
        "birth_date",
        "system_status",
        "hr_status",
        "branch_basis",
        "parent_branch",
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

    if "department" in system.columns:
        system["department"] = system["department"].astype(str).fillna("").str.strip()
    else:
        system["department"] = ""

    branch_column = resolve_system_branch_column(system)
    if branch_column and branch_column in system.columns:
        system["branch_basis"] = system[branch_column].astype(str).fillna("").str.strip()
    else:
        system["branch_basis"] = ""
    system["parent_branch"] = system["branch_basis"].map(map_system_branch_label)
    system = build_full_name(system)

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
    system_report = system[["id", "system_status", "department", "branch_basis", "parent_branch"]].copy()
    matched_employees = system_report.merge(hr, on="id", how="inner")

    inactive_to_update = matched_employees[
        (matched_employees["system_status"] == "active") &
        (matched_employees["hr_status"] != "active")
    ].copy()
    inactive_to_update["hr_status"] = "inactive"

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
    run_branch_key = determine_run_branch_key(system, hr_path)
    reports = generate_reports(system, hr)
    new_in_system_report, new_in_system_meta = build_new_in_system_report_for_branch(system, hr, run_branch_key)

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
    new_in_system_output = save_report(
        new_in_system_report,
        REPORT_FILENAMES["new_in_system_since_last_month"],
        output_dir,
    )
    branch_system = system[system["parent_branch"] == run_branch_key].copy()
    snapshot_dirs = []
    if not branch_system.empty:
        snapshot_dirs.append(save_branch_snapshot(run_branch_key, system_path, branch_system, output_dir, hr_path))
    return {
        "system_file": system_path.name,
        "hr_file": hr_path.name,
        "inactive_output": inactive_output,
        "active_output": active_output,
        "new_active_output": new_active_output,
        "new_in_system_output": new_in_system_output,
        "matched_count": int(reports["_matched_count"].iloc[0]["count"]),
        "inactive_count": len(reports["inactive_to_update"]),
        "active_count": len(reports["active_to_update"]),
        "new_active_count": len(reports["new_active_employees"]),
        "new_in_system_count": int(new_in_system_meta["count"]),
        "new_in_system_active_count": int(new_in_system_meta["active_count"]),
        "new_in_system_inactive_count": int(new_in_system_meta["inactive_count"]),
        "new_in_system_previous_snapshot_found": bool(new_in_system_meta["previous_snapshot_found"]),
        "new_in_system_previous_snapshot_name": str(new_in_system_meta["previous_snapshot_name"]),
        "new_in_system_missing_snapshot_branches": list(new_in_system_meta["missing_snapshot_branches"]),
        "run_branch_key": run_branch_key,
        "snapshot_dirs": snapshot_dirs,
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
    print(f"New in system report: {Path(result['new_in_system_output']).name}")
    print(f"Matched employees checked: {result['matched_count']}")
    print(f"Set to inactive: {result['inactive_count']}")
    print(f"Set to active: {result['active_count']}")
    print(f"New active employees not in system: {result['new_active_count']}")
    print(f"New in system since last snapshot: {result['new_in_system_count']}")


if __name__ == "__main__":
    main()

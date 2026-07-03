"""Combine per-owner commission workbooks into a single workbook.

Walks a root folder (e.g. ``output_by_manager``) for ``*.xlsx`` files and copies
every sheet into one combined workbook, giving each sheet a descriptive name.

A sheet is classified by its **sheet name**, not its position:
- a sheet named ``Manager_Rollup`` is a managerial commission -> ``*_managerial_com``
- any other sheet is an individual commission            -> ``*_individual_com``

This matters because a workbook may contain a single ``Manager_Rollup`` sheet
(a manager with no personal sales, e.g. Jason Jones) or both an individual sheet
and a ``Manager_Rollup`` sheet (a manager who also sells, e.g. Dirk Edwards).
"""

from pathlib import Path

import pandas as pd

# Sheet name written by test_comm.py for managerial rollups.
MANAGER_ROLLUP_SHEET = "Manager_Rollup"

# Column in individual sheets holding the sales rep.
OWNER_COLUMN = "Opportunity Owner"
# Column in managerial rollups holding the manager the rollup belongs to.
MANAGER_COLUMN = "Mapped_Name"

# Excel caps sheet names at 31 characters.
MAX_SHEET_NAME = 31


def clean_name(name: object) -> str:
    """Lowercase, strip, and replace spaces with underscores."""
    return str(name).strip().lower().replace(" ", "_")


def get_owner_name(df: pd.DataFrame, column: str) -> str:
    """Return the first valid owner name from an individual sheet."""
    if column not in df.columns:
        return "unknown_owner"

    owners = df[column].dropna().unique()
    if len(owners) == 0:
        return "unknown_owner"

    return clean_name(owners[0])


def get_manager_name(df: pd.DataFrame, file_path: Path) -> str:
    """Return the manager a rollup belongs to.

    Prefers the ``Mapped_Name`` column; falls back to the file name so a
    managerial-only workbook is still labelled after the right person.
    """
    if MANAGER_COLUMN in df.columns:
        managers = df[MANAGER_COLUMN].dropna().unique()
        if len(managers) > 0:
            return clean_name(managers[0])

    return clean_name(file_path.stem)


def unique_sheet_name(name: str, used: set[str]) -> str:
    """Truncate to Excel's limit and de-duplicate against already-used names."""
    candidate = name[:MAX_SHEET_NAME]
    counter = 1
    while candidate in used:
        suffix = f"_{counter}"
        candidate = f"{name[:MAX_SHEET_NAME - len(suffix)]}{suffix}"
        counter += 1
    used.add(candidate)
    return candidate


def combine_workbooks(root_dir: Path, output_file: Path) -> int:
    """Combine every sheet under ``root_dir`` into ``output_file``.

    Returns:
        The number of sheets written.
    """
    excel_files = sorted(root_dir.rglob("*.xlsx"))
    print(f"Found {len(excel_files)} Excel files")

    used_names: set[str] = set()
    sheet_counter = 0

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for file_path in excel_files:
            try:
                excel = pd.ExcelFile(file_path)
                for sheet in excel.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet)

                    if sheet == MANAGER_ROLLUP_SHEET:
                        base_name = get_manager_name(df, file_path)
                        label = "managerial_com"
                    else:
                        base_name = get_owner_name(df, OWNER_COLUMN)
                        if base_name == "unknown_owner":
                            base_name = clean_name(file_path.stem)
                        label = "individual_com"

                    sheet_name = unique_sheet_name(
                        f"{base_name}_{label}", used_names
                    )
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    sheet_counter += 1
            except Exception as error:
                print(f"Error processing {file_path}: {error}")

    print(f"Processed {sheet_counter} sheets")
    return sheet_counter


def main() -> None:
    """Prompt for the input folder and write the combined workbook."""
    root_dir = Path(input("Enter the folder to combine (e.g. output_by_manager): ").strip()).expanduser()
    if not root_dir.exists():
        raise FileNotFoundError(f"Folder not found: {root_dir}")

    output_file = Path("combined_output.xlsx")
    combine_workbooks(root_dir, output_file)
    print(f"Workbook saved as {output_file}")


if __name__ == "__main__":
    main()

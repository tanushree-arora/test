from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import pandas as pd

import warnings

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style",
    category=UserWarning,
    module="openpyxl.styles.stylesheet",
)

MAIN_SKIPROWS = 13
FX_SKIPROWS = 15

OWNER_REQUIRED_COLUMNS = ["Full Name", "Manager", "Team"]
MAIN_REQUIRED_COLUMNS = ["Opportunity Owner", "Type", "Product Type"]
MAIN_SELECTED_COLUMNS = [
    "Opportunity Owner",
    "Account Name",
    "Opportunity Name",
    "Product Name",
    "Product Type",
    "Sale Type",
    "Billing Date",
    "Close Date",
    "Opportunity Currency",
    "Total Line Value",
    "Brand",
    "Type",
]
COMM_REQUIRED_COLUMNS = [
    "Role",
    "Name",
    "Sale Type",
    "Product Type",
    "Customer Type",
    "Brand",
    "Customer",
]
FX_REQUIRED_COLUMNS = ["Currency Code", "END OF MONTH"]

FINAL_COMMISSION_COLUMNS = [
    "Opportunity Owner",
    "Account Name",
    "Opportunity Name",
    "Product Name",
    "Product Type",
    "Sale Type",
    "Billing Date",
    "Close Date",
    "Opportunity Currency",
    "Total Line Value",
    "Brand",
    "Type",
    "Manager",
    "Team",
    "Commission Rate",
    "FX rate",
    "Location FX Factor",
    "Location Currency",
    "Final Commission",
    "Final Commission Adjusted",
]

OWNER_CURRENCY_MAP = {
    "dirk edwards": "CAD",
    "thalia kouts": "AUD",
}

MANAGER_TEAMS = {
    "jason jones": [
        "ben braman",
        "jake mickelson",
        "john turner",
        "jordan stuart",
        "justice forbes",
        "mitch ryan",
        "scott carlyle",
        "taylor dennis",
        "ty walker",
    ],
    "dirk edwards": [
        "angela burciaga",
        "cassandra fortner",
        "courtney kocel",
        "jessica wheelin",
        "kaylen everhart",
        "kevin williams",
        "michelle spencer",
        "tiffany elliott",
    ],
}


def prompt_path(prompt_text: str) -> Path:
    """Prompt for a file path and validate it exists.

    Args:
        prompt_text: Prompt text shown to the user.

    Returns:
        A validated Path object.

    Raises:
        FileNotFoundError: If the entered file does not exist.
    """
    path = Path(input(prompt_text).strip()).expanduser()
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    return path


def normalize_series(series: pd.Series, uppercase: bool = False) -> pd.Series:
    """Normalize text values for reliable joins.

    Args:
        series: Input Series.
        uppercase: Whether to convert to uppercase instead of lowercase.

    Returns:
        A normalized Series.
    """
    normalized = series.astype(str).str.strip()
    normalized = normalized.str.upper() if uppercase else normalized.str.lower()
    return normalized.replace({"nan": None, "none": None, "": None})


def validate_columns(
    dataframe: pd.DataFrame,
    required_columns: Iterable[str],
    dataframe_name: str,
) -> None:
    """Ensure the DataFrame contains the required columns.

    Args:
        dataframe: DataFrame to validate.
        required_columns: Expected column names.
        dataframe_name: Friendly name used in error messages.

    Raises:
        ValueError: If any required columns are missing.
    """
    missing = [column for column in required_columns if column not in dataframe.columns]
    if missing:
        raise ValueError(
            f"{dataframe_name} is missing required columns: {missing}"
        )


def read_input_files() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Read all required Excel input files.

    Returns:
        Tuple containing main report, owner mapping, commission mapping, and FX table.
    """
    main_path = prompt_path("Enter the main commission report Excel path: ")
    owner_path = prompt_path("Enter the owner mapping Excel path: ")
    commission_path = prompt_path("Enter the commission mapping Excel path: ")
    fx_path = prompt_path("Enter the FX Excel path: ")

    df_main = pd.read_excel(main_path, skiprows=MAIN_SKIPROWS)
    df_owner = pd.read_excel(owner_path)
    df_comm = pd.read_excel(commission_path)
    fx_table = pd.read_excel(fx_path, skiprows=FX_SKIPROWS)[
        ["Currency Code", "END OF MONTH"]
    ]

    for dataframe in (df_main, df_owner, df_comm, fx_table):
        dataframe.columns = dataframe.columns.str.strip()

    return df_main, df_owner, df_comm, fx_table


def merge_owner_mapping(
    df_main: pd.DataFrame,
    df_owner: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.Series]:
    """Merge owner-manager-team mapping into the main dataset.

    Args:
        df_main: Main report DataFrame.
        df_owner: Owner mapping DataFrame.

    Returns:
        Tuple of merged DataFrame and unmatched owner Series.
    """
    validate_columns(df_main, MAIN_REQUIRED_COLUMNS, "Main dataset")
    validate_columns(df_owner, OWNER_REQUIRED_COLUMNS, "Owner mapping dataset")
    validate_columns(df_main, MAIN_SELECTED_COLUMNS, "Main dataset")

    main_filtered = df_main[MAIN_SELECTED_COLUMNS].copy()
    owner_filtered = df_owner[OWNER_REQUIRED_COLUMNS].copy()

    main_filtered["Opportunity Owner_clean"] = normalize_series(
        main_filtered["Opportunity Owner"]
    )
    owner_filtered["Full Name_clean"] = normalize_series(owner_filtered["Full Name"])
    owner_filtered = owner_filtered.drop_duplicates(subset=["Full Name_clean"])

    merged = main_filtered.merge(
        owner_filtered,
        how="left",
        left_on="Opportunity Owner_clean",
        right_on="Full Name_clean",
    )

    unmatched_owners = (
        merged.loc[merged["Manager"].isna(), "Opportunity Owner"]
        .drop_duplicates()
   #     .sort_values()
    )

    return merged, unmatched_owners


def add_commission_join_keys(
    df_step1: pd.DataFrame,
    df_comm: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Add normalized join keys to the sales and commission datasets.

    Args:
        df_step1: Sales DataFrame after owner merge.
        df_comm: Commission mapping DataFrame.

    Returns:
        Tuple of transformed sales and commission DataFrames.
    """
    validate_columns(
        df_step1,
        [
            "Team",
            "Opportunity Owner",
            "Sale Type",
            "Product Type",
            "Type",
            "Brand",
            "Account Name",
        ],
        "Step 1 dataset",
    )
    validate_columns(df_comm, COMM_REQUIRED_COLUMNS, "Commission dataset")

    step1 = df_step1.copy()
    comm = df_comm.copy()

    step1["Team_clean"] = normalize_series(step1["Team"])
    step1["Opportunity Owner_clean"] = normalize_series(step1["Opportunity Owner"])
    step1["Sale Type_clean"] = normalize_series(step1["Sale Type"])
    step1["Product Type_clean"] = normalize_series(step1["Product Type"])
    step1["Type_clean"] = normalize_series(step1["Type"])
    step1["Brand_clean"] = normalize_series(step1["Brand"])
    step1["Account Name_clean"] = normalize_series(step1["Account Name"])

    comm["Role_clean"] = normalize_series(comm["Role"])
    comm["Name_clean"] = normalize_series(comm["Name"])
    comm["Sales Type_clean"] = normalize_series(comm["Sale Type"])
    comm["Product Type_clean"] = normalize_series(comm["Product Type"])
    comm["Customer Type_clean"] = normalize_series(comm["Customer Type"])
    comm["Brand_clean"] = normalize_series(comm["Brand"])
    comm["Customer_clean"] = normalize_series(comm["Customer"])

    return step1, comm


def apply_conditional_commission_merge(
    df_step1: pd.DataFrame,
    df_comm: pd.DataFrame,
) -> pd.DataFrame:
    """Apply business-rule-specific commission merges.

    Args:
        df_step1: Sales DataFrame after owner merge.
        df_comm: Commission mapping DataFrame.

    Returns:
        Merged DataFrame with commission rule matches.
    """
    step1, comm = add_commission_join_keys(df_step1, df_comm)

    mask_ronan = step1["Opportunity Owner_clean"] == "ronan o'maitiu"
    mask_alejandro = step1["Opportunity Owner_clean"] == "alejandro varela"
    mask_account_exec = (
        (step1["Team_clean"] == "account executive")
        & ~mask_ronan
        & ~mask_alejandro
    )
    mask_default = ~(mask_ronan | mask_alejandro | mask_account_exec)

    df_default = step1.loc[mask_default].merge(
        comm,
        how="left",
        left_on=[
            "Team_clean",
            "Opportunity Owner_clean",
            "Sale Type_clean",
            "Product Type_clean",
        ],
        right_on=[
            "Role_clean",
            "Name_clean",
            "Sales Type_clean",
            "Product Type_clean",
        ],
        suffixes=("", "_comm"),
    )

    df_account_exec = step1.loc[mask_account_exec].merge(
        comm,
        how="left",
        left_on=[
            "Team_clean",
            "Opportunity Owner_clean",
            "Sale Type_clean",
            "Product Type_clean",
            "Type_clean",
        ],
        right_on=[
            "Role_clean",
            "Name_clean",
            "Sales Type_clean",
            "Product Type_clean",
            "Customer Type_clean",
        ],
        suffixes=("", "_comm"),
    )

    df_alejandro = step1.loc[mask_alejandro].merge(
        comm,
        how="left",
        left_on=[
            "Team_clean",
            "Opportunity Owner_clean",
            "Sale Type_clean",
            "Product Type_clean",
            "Brand_clean",
        ],
        right_on=[
            "Role_clean",
            "Name_clean",
            "Sales Type_clean",
            "Product Type_clean",
            "Brand_clean",
        ],
        suffixes=("", "_comm"),
    )

    df_ronan = step1.loc[mask_ronan].merge(
        comm,
        how="left",
        left_on=[
            "Team_clean",
            "Opportunity Owner_clean",
            "Sale Type_clean",
            "Product Type_clean",
            "Account Name_clean",
        ],
        right_on=[
            "Role_clean",
            "Name_clean",
            "Sales Type_clean",
            "Product Type_clean",
            "Customer_clean",
        ],
        suffixes=("", "_comm"),
    )

    merged = pd.concat(
        [df_default, df_account_exec, df_alejandro, df_ronan],
        ignore_index=True,
    )

    helper_columns = [
        "Team_clean",
        "Opportunity Owner_clean",
        "Sale Type_clean",
        "Product Type_clean",
        "Type_clean",
        "Brand_clean",
        "Account Name_clean",
        "Role_clean",
        "Name_clean",
        "Sales Type_clean",
        "Customer Type_clean",
        "Customer_clean",
    ]
    return merged.drop(columns=helper_columns, errors="ignore")


def calculate_sales_commission(
    df: pd.DataFrame,
    fx_table: pd.DataFrame,
) -> pd.DataFrame:
    """Calculate commission, FX-adjusted commission, and location-adjusted payout.

    Note:
        This preserves your original behavior for Rate handling. If Rate values
        like 5% should mean 0.05, then divide by 100 before calculating.

    Args:
        df: Commission-matched sales DataFrame.
        fx_table: FX lookup table.

    Returns:
        Final commission DataFrame.
    """
    working_df = df.copy()
    fx_lookup = fx_table.copy()

    validate_columns(working_df, ["Rate", "Total Line Value"], "Commission dataset")
    validate_columns(fx_lookup, FX_REQUIRED_COLUMNS, "FX dataset")

    working_df["Location Currency"] = normalize_series(
        working_df["Opportunity Owner"]
    ).map(OWNER_CURRENCY_MAP).fillna("USD")

    working_df = working_df.rename(columns={"Rate": "Commission Rate"})
    working_df["Commission Rate"] = (
        working_df["Commission Rate"]
        .astype(str)
        .str.strip()
        .str.replace("%", "", regex=False)
    )
    working_df["Commission Rate"] = pd.to_numeric(
        working_df["Commission Rate"],
        errors="coerce",
    )
    working_df["Total Line Value"] = pd.to_numeric(
        working_df["Total Line Value"],
        errors="coerce",
    )
    working_df["Commission"] = (
        working_df["Commission Rate"] * working_df["Total Line Value"]
    )

    fx_lookup["Currency Code_clean"] = normalize_series(
        fx_lookup["Currency Code"],
        uppercase=True,
    )
    working_df["Opportunity Currency_clean"] = normalize_series(
        working_df["Opportunity Currency"],
        uppercase=True,
    )

    working_df = working_df.merge(
        fx_lookup[["Currency Code_clean", "END OF MONTH"]],
        how="left",
        left_on="Opportunity Currency_clean",
        right_on="Currency Code_clean",
    )

    usd_mask = working_df["Opportunity Currency_clean"] == "USD"
    working_df.loc[usd_mask, "END OF MONTH"] = 1
    working_df = working_df.rename(columns={"END OF MONTH": "FX rate"})

    working_df["Commission"] = pd.to_numeric(
        working_df["Commission"],
        errors="coerce",
    )
    working_df["FX rate"] = pd.to_numeric(
        working_df["FX rate"],
        errors="coerce",
    )
    working_df["Final Commission"] = (
        working_df["Commission"] * working_df["FX rate"]
    ).fillna(0).round(2)

    location_fx_lookup = fx_lookup[
        ["Currency Code_clean", "END OF MONTH"]
    ].drop_duplicates(subset=["Currency Code_clean"])

    working_df["Location Currency_clean"] = normalize_series(
        working_df["Location Currency"],
        uppercase=True,
    )

    working_df = working_df.merge(
        location_fx_lookup,
        how="left",
        left_on="Location Currency_clean",
        right_on="Currency Code_clean",
        suffixes=("", "_location"),
    )

    working_df["END OF MONTH"] = pd.to_numeric(
        working_df["END OF MONTH"],
        errors="coerce",
    )
    working_df["Location FX Factor"] = 1.0

    non_usd_mask = working_df["Location Currency_clean"] != "USD"
    working_df.loc[non_usd_mask, "Location FX Factor"] = (
        1 / working_df.loc[non_usd_mask, "END OF MONTH"]
    )

    working_df["Final Commission Adjusted"] = (
        working_df["Final Commission"] * working_df["Location FX Factor"]
    ).round(2)

    # Presentation formatting added here.
    working_df["Commission Rate"] = working_df["Commission Rate"].apply(
        lambda value: f"{value* 100:.4f}%" if pd.notnull(value) else ""
    )

    currency_cols = ["Final Commission Adjusted"]
    for column in currency_cols:
        if column in working_df.columns:
            working_df[column] = working_df[column].apply(
                lambda value: f"${value:,.2f}" if pd.notnull(value) else ""
            )

    available_columns = [
        column for column in FINAL_COMMISSION_COLUMNS
        if column in working_df.columns
    ]
    return working_df[available_columns].reset_index(drop=True)


def make_safe_name(value: object) -> str:
    """Convert a value into a filesystem-safe file or folder name.

    Args:
        value: Any input value.

    Returns:
        Sanitized name safe for filesystem use.
    """
    safe_value = str(value).strip()
    safe_value = re.sub(r'[\\/*?:"<>|]', "_", safe_value)
    return re.sub(r"\s+", "_", safe_value)


def export_owner_reports(
    df_final: pd.DataFrame,
    output_dir: Path,
    unmatched_owners: pd.Series,
) -> list[Path]:
    """Export one file per manager and opportunity owner.

    Args:
        df_final: Final commission DataFrame.
        output_dir: Base output folder.
        unmatched_owners: Owners not found in owner mapping.

    Returns:
        List of saved report paths.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    saved_files: list[Path] = []

    for (manager, owner), group in df_final.groupby(["Manager", "Opportunity Owner"]):
        manager_folder = output_dir / make_safe_name(manager)
        manager_folder.mkdir(parents=True, exist_ok=True)

        output_path = manager_folder / f"{make_safe_name(owner)}.xlsx"
        group.reset_index(drop=True).to_excel(output_path, index=False)
        saved_files.append(output_path)

    if not unmatched_owners.empty:
        pd.DataFrame(
            {"Unmatched Opportunity Owner": unmatched_owners}
        ).to_excel(output_dir / "unmatched_owners.xlsx", index=False)

    return saved_files


def build_managerial_commission_report(
    df_final: pd.DataFrame,
    df_comm: pd.DataFrame,
) -> pd.DataFrame:
    """Build the managerial commission report.

    Args:
        df_final: Final sales commission DataFrame.
        df_comm: Original commission rule DataFrame.

    Returns:
        Managerial commission report DataFrame.
    """
    owner_to_manager = {
        owner: manager
        for manager, members in MANAGER_TEAMS.items()
        for owner in members
    }

    final_working = df_final.copy()
    comm_working = df_comm.copy()

    for column in ["Opportunity Owner", "Sale Type", "Product Type", "Brand"]:
        if column in final_working.columns:
            final_working[column] = normalize_series(final_working[column])

    for column in ["Name", "Sale Type", "Product Type", "Brand"]:
        if column in comm_working.columns:
            comm_working[column] = normalize_series(comm_working[column])

    final_working["Mapped_Name"] = final_working["Opportunity Owner"].map(
        owner_to_manager
    )

    filtered_final = final_working.loc[
        ~final_working["Mapped_Name"].isna(),
        [
            "Opportunity Owner",
            "Product Name",
            "Product Type",
            "Sale Type",
            "Billing Date",
            "Close Date",
            "Opportunity Currency",
            "Total Line Value",
            "Brand",
            "Type",
            "FX rate",
            "Location FX Factor",
            "Location Currency",
            "Mapped_Name",
        ],
    ].copy()

    manager_rules = comm_working.loc[
        comm_working["Role"].eq("Manager"),
        ["Name", "Brand", "Sale Type", "Product Type", "Rate"],
    ].drop_duplicates()

    df_jason = filtered_final[filtered_final["Mapped_Name"] == "jason jones"]
    df_dirk = filtered_final[filtered_final["Mapped_Name"] == "dirk edwards"]

    rules_jason = manager_rules[manager_rules["Name"] == "jason jones"]
    rules_dirk = manager_rules[manager_rules["Name"] == "dirk edwards"]

    merged_jason = df_jason.merge(
        rules_jason,
        how="left",
        left_on=["Mapped_Name", "Sale Type", "Product Type", "Brand"],
        right_on=["Name", "Sale Type", "Product Type", "Brand"],
    )

    merged_dirk = df_dirk.merge(
        rules_dirk,
        how="left",
        left_on=["Mapped_Name", "Sale Type", "Product Type"],
        right_on=["Name", "Sale Type", "Product Type"],
    )

    result = pd.concat([merged_jason, merged_dirk], ignore_index=True)

    result = result[
        [
            "Opportunity Owner",
            "Product Name",
            "Product Type",
            "Sale Type",
            "Billing Date",
            "Close Date",
            "Opportunity Currency",
            "Total Line Value",
            "Brand",
            "Type",
            "FX rate",
            "Location FX Factor",
            "Location Currency",
            "Mapped_Name",
            "Rate",
        ]
    ].rename(columns={"Rate": "Managerial comm rate"})

    result["Total Line Value"] = pd.to_numeric(
        result["Total Line Value"],
        errors="coerce",
    )
    result["Managerial comm rate"] = pd.to_numeric(
        result["Managerial comm rate"],
        errors="coerce",
    )
    result["FX rate"] = pd.to_numeric(result["FX rate"], errors="coerce")
    result["Location FX Factor"] = pd.to_numeric(
        result["Location FX Factor"],
        errors="coerce",
    )

    result["Managerial Commission"] = (
        result["Total Line Value"]
        * result["Managerial comm rate"]
        * result["FX rate"]
        * result["Location FX Factor"]
    )

    result["Managerial comm rate"] = result["Managerial comm rate"].apply(
        lambda value: f"{value* 100:.4f}%" if pd.notnull(value) else ""
    )
    result["Managerial Commission"] = result["Managerial Commission"].apply(
        lambda value: f"${value:,.2f}" if pd.notnull(value) else ""
    )
    result["Mapped_Name"] = result["Mapped_Name"].str.title()

    return result


def export_manager_reports(
    df_result: pd.DataFrame,
    output_dir: Path,
    parent_lookup: dict[str, str],
) -> None:
    """Export one workbook per manager, nested under their parent manager.

    Each managerial rollup is placed inside a folder named after the manager
    that Dirk/Jason report to, mirroring the owner report layout so the merge
    step can pair same-named workbooks across both trees.

    Args:
        df_result: Managerial commission report.
        output_dir: Destination folder.
        parent_lookup: Mapping of normalized manager name -> parent manager name.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    for manager, group in df_result.groupby("Mapped_Name"):
        parent = parent_lookup.get(str(manager).strip().lower())
        if pd.notnull(parent) and str(parent).strip():
            parent_folder_name = make_safe_name(parent)
        else:
            parent_folder_name = "Unassigned"

        parent_folder = output_dir / parent_folder_name
        parent_folder.mkdir(parents=True, exist_ok=True)

        output_path = parent_folder / f"{make_safe_name(manager)}.xlsx"
        group.to_excel(output_path, index=False)


def merge_duplicate_workbooks(
    root_folders: list[Path],
    delete_duplicate_after_merge: bool = False,
) -> None:
    """Merge workbooks that share the same filename across folders.

    Args:
        root_folders: Folders to search recursively for Excel files.
        delete_duplicate_after_merge: Whether to delete merged duplicate files.
    """
    all_files = [
        file_path
        for folder in root_folders
        for file_path in folder.rglob("*.xlsx")
        if file_path.is_file()
    ]

    files_by_name: dict[str, list[Path]] = {}
    for file_path in all_files:
        files_by_name.setdefault(file_path.name, []).append(file_path)

    duplicate_files = {
        file_name: paths
        for file_name, paths in files_by_name.items()
        if len(paths) > 1
    }

    for _, paths in duplicate_files.items():
        primary_file = paths[0]
        duplicates = paths[1:]

        for duplicate_index, second_file in enumerate(duplicates, start=2):
            second_excel = pd.ExcelFile(second_file)

            with pd.ExcelWriter(
                primary_file,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="new",
            ) as writer:
                existing_sheet_names = set(writer.book.sheetnames)

                for sheet_name in second_excel.sheet_names:
                    temp_df = pd.read_excel(second_file, sheet_name=sheet_name)

                    if len(second_excel.sheet_names) > 1:
                        new_sheet_name = f"{second_file.stem}_{sheet_name}"[:31]
                    else:
                        new_sheet_name = f"Sheet{duplicate_index}"[:31]

                    original_name = new_sheet_name
                    counter = 1
                    while new_sheet_name in existing_sheet_names:
                        suffix = f"_{counter}"
                        new_sheet_name = (
                            f"{original_name[:31 - len(suffix)]}{suffix}"
                        )
                        counter += 1

                    temp_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
                    existing_sheet_names.add(new_sheet_name)

            if delete_duplicate_after_merge:
                second_file.unlink(missing_ok=True)


def main() -> None:
    """Run the full commission processing pipeline."""
    df_main, df_owner, df_comm, fx_table = read_input_files()

    owner_merged, unmatched_owners = merge_owner_mapping(df_main, df_owner)
    commission_merged = apply_conditional_commission_merge(owner_merged, df_comm)
    df_final = calculate_sales_commission(commission_merged, fx_table)

    owner_output_dir = Path("output_by_manager")
    manager_output_dir = Path("manager_reports")

    saved_files = export_owner_reports(
        df_final=df_final,
        output_dir=owner_output_dir,
        unmatched_owners=unmatched_owners,
    )
    print(f"Created {len(saved_files)} owner files successfully.")

    summary = (
        df_final.groupby(["Manager", "Team", "Opportunity Owner"])
        .size()
        .reset_index(name="record_count")
        .sort_values(["Manager", "Team", "Opportunity Owner"])
    )
    print(summary.head())

    df_manager_report = build_managerial_commission_report(df_final, df_comm)

    parent_lookup = dict(
        zip(normalize_series(df_owner["Full Name"]), df_owner["Manager"])
    )
    export_manager_reports(df_manager_report, manager_output_dir, parent_lookup)
    print("Manager report files created successfully.")

    merge_duplicate_workbooks(
        root_folders=[owner_output_dir, manager_output_dir],
        delete_duplicate_after_merge=False,
    )
    print("Duplicate workbook merge completed.")


if __name__ == "__main__":
    main()
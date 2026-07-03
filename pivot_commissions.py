"""Build a unified commissions pivot from combined_output.xlsx.

Combines individual commissions (``Final Commission Adjusted``) and managerial
commissions (``Managerial Commission``) into one pivot:

- Rows:    Person (the commission recipient)
- Columns: Sale Type -> Brand -> Product Type
- Values:  total commission (individual + managerial), summed
- Row totals per person and a grand total (``margins``)

Each sheet in combined_output.xlsx is one person's commissions. A sheet is an
individual sheet unless it carries a ``Managerial Commission`` column, in which
case the recipient is the manager (``Mapped_Name``) rather than the seller.
"""

from pathlib import Path

import pandas as pd

INDIVIDUAL_VALUE = "Final Commission Adjusted"
MANAGERIAL_VALUE = "Managerial Commission"

CATEGORY_COLUMNS = ["Sale Type", "Brand", "Product Type"]


def to_number(series: pd.Series) -> pd.Series:
    """Parse a ``$1,234.56`` currency-string column back into floats."""
    cleaned = (
        series.astype(str)
        .str.replace(r"[$,]", "", regex=True)
        .str.strip()
        .replace({"": None, "nan": None, "None": None})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def normalize_label(series: pd.Series) -> pd.Series:
    """Title-case category/person text so both sources share the same labels."""
    return (
        series.astype(str)
        .str.strip()
        .str.title()
        .replace({"Nan": "", "None": ""})
    )


def column_or_blank(df: pd.DataFrame, column: str) -> pd.Series:
    """Return the column if present, else an all-blank Series of the right length."""
    if column in df.columns:
        return df[column]
    return pd.Series([""] * len(df), index=df.index)


def load_records(input_file: Path) -> pd.DataFrame:
    """Read every sheet and return long-format rows: Person + categories + amount."""
    records: list[pd.DataFrame] = []

    excel = pd.ExcelFile(input_file)
    for sheet in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet)

        if MANAGERIAL_VALUE in df.columns:
            person_col, value_col = "Mapped_Name", MANAGERIAL_VALUE
        elif INDIVIDUAL_VALUE in df.columns:
            person_col, value_col = "Opportunity Owner", INDIVIDUAL_VALUE
        else:
            # e.g. unmatched_owners sheet - nothing to sum, skip it.
            continue

        part = pd.DataFrame(
            {
                "Person": normalize_label(column_or_blank(df, person_col)),
                "Sale Type": normalize_label(column_or_blank(df, "Sale Type")),
                "Brand": normalize_label(column_or_blank(df, "Brand")),
                "Product Type": normalize_label(column_or_blank(df, "Product Type")),
                "Commission": to_number(df[value_col]),
            }
        )
        records.append(part)

    if not records:
        raise ValueError(f"No commission sheets found in {input_file}")

    data = pd.concat(records, ignore_index=True)
    data["Commission"] = data["Commission"].fillna(0.0)
    return data


def build_pivot(data: pd.DataFrame) -> pd.DataFrame:
    """Pivot: Person x (Sale Type, Brand, Product Type), summed, with totals."""
    return pd.pivot_table(
        data,
        index="Person",
        columns=CATEGORY_COLUMNS,
        values="Commission",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Grand Total",
    ).round(2)


def main() -> None:
    """Read combined_output.xlsx and write the unified commissions pivot."""
    entered = input(
        "Enter the combined workbook path [combined_output.xlsx]: "
    ).strip()
    input_file = Path(entered) if entered else Path("combined_output.xlsx")
    if not input_file.exists():
        raise FileNotFoundError(f"File not found: {input_file}")

    output_file = Path("commissions_pivot.xlsx")

    data = load_records(input_file)
    pivot = build_pivot(data)
    pivot.to_excel(output_file)

    print(f"Combined {len(data)} rows across {data['Person'].nunique()} people.")
    print(f"Pivot shape (rows x cols): {pivot.shape}")
    print(f"Saved {output_file}")


if __name__ == "__main__":
    main()

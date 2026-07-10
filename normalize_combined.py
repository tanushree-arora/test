"""Flatten combined_output.xlsx into a single normalized sheet.

Reads every sheet, classifies it as individual or managerial (by the presence
of a ``Managerial Commission`` column), keeps all detail columns, and stacks
everything into one sheet with a consistent ``Person``, ``Commission``, and
``Commission Type`` column so any downstream tool can sum ``Commission``
without double-counting or misidentifying columns.
"""

from pathlib import Path

import pandas as pd

INDIVIDUAL_VALUE = "Final Commission Adjusted"
MANAGERIAL_VALUE = "Managerial Commission"


def normalize_combined(input_file: Path, output_file: Path) -> None:
    """Read combined_output.xlsx and write a single-sheet normalized version."""
    excel = pd.ExcelFile(input_file)
    records: list[pd.DataFrame] = []

    for sheet in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet)

        if MANAGERIAL_VALUE in df.columns:
            df = df.rename(columns={
                "Mapped_Name": "Person",
                MANAGERIAL_VALUE: "Commission",
            })
            df["Commission Type"] = "Managerial"
        elif INDIVIDUAL_VALUE in df.columns:
            df = df.rename(columns={
                "Opportunity Owner": "Person",
                INDIVIDUAL_VALUE: "Commission",
            })
            df["Commission Type"] = "Individual"
        else:
            continue

        df["Source Sheet"] = sheet
        records.append(df)

    if not records:
        raise ValueError(f"No commission sheets found in {input_file}")

    combined = pd.concat(records, ignore_index=True)

    cols = ["Person", "Commission", "Commission Type", "Source Sheet"]
    rest = [c for c in combined.columns if c not in cols]
    combined = combined[cols + rest]

    combined.to_excel(output_file, index=False, sheet_name="Normalized")
    people = combined["Person"].nunique()
    total = combined["Commission"].sum()
    print(f"{len(combined)} rows, {people} people, total commission: ${total:,.2f}")
    print(f"Saved {output_file}")


def main() -> None:
    entered = input(
        "Enter the combined workbook path [combined_output.xlsx]: "
    ).strip()
    input_file = Path(entered) if entered else Path("combined_output.xlsx")
    if not input_file.exists():
        raise FileNotFoundError(f"File not found: {input_file}")

    output_file = Path("combined_normalized.xlsx")
    normalize_combined(input_file, output_file)


if __name__ == "__main__":
    main()

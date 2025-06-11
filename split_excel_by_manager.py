import os
import re
import pandas as pd
import argparse


def sanitize_filename(name: str) -> str:
    """Return a filesystem-safe filename from a string."""
    sanitized = re.sub(r'[\\/*?:"<>|]', '_', str(name))
    sanitized = sanitized.strip()
    return sanitized or "Unknown_Manager"


def split_excel_by_manager(input_path: str, output_dir: str, manager_column: str = "Manager") -> None:
    """Split an Excel sheet into multiple files based on the manager column."""
    df = pd.read_excel(input_path)
    if manager_column not in df.columns:
        raise ValueError(f"Column '{manager_column}' not found in input file")

    grouped = df.groupby(manager_column)
    os.makedirs(output_dir, exist_ok=True)

    for manager, group in grouped:
        manager_safe = sanitize_filename(manager)
        output_file = os.path.join(output_dir, f"{manager_safe}.xlsx")
        group.to_excel(output_file, index=False)

    print(f"Saved {len(grouped)} file(s) to '{output_dir}'.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Split an Excel file by the Manager column.")
    parser.add_argument("input_excel", help="Path to the source Excel file")
    parser.add_argument("-o", "--output-dir", default="output", help="Directory for the split files")
    parser.add_argument("-c", "--column", default="Manager", help="Column name that contains manager names")
    args = parser.parse_args()

    split_excel_by_manager(args.input_excel, args.output_dir, args.column)


if __name__ == "__main__":
    main()

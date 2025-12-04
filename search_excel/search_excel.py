import os
import argparse
import pandas as pd

def search_excel_with_pandas(folder_path, search_text):
    search_text = search_text.lower()
    results = []

    for filename in os.listdir(folder_path):
        if not filename.lower().endswith((".xlsx", ".xls", ".xlsm")):
            continue

        file_path = os.path.join(folder_path, filename)
        print(f"Searching {filename}...")

        try:
            sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        except Exception as e:
            print(f"Could not read {filename}: {e}")
            continue

        for sheet_name, df in sheets.items():
            df = df.fillna("")

            # True/False mask of where matches occur
            mask = df.apply(lambda col: col.str.lower().str.contains(search_text, na=False))

            # Rows/columns that matched
            match_positions = mask.stack()[mask.stack()].index

            for row, col in match_positions:
                value = df.at[row, col]
                results.append({
                    "file": filename,
                    "sheet": sheet_name,
                    "row": row + 2,   # Excel-style row (assuming header on row 1)
                    "column": col,
                    "value": value
                })

    return results


def main():
    parser = argparse.ArgumentParser(
        description="Search Excel files for a partial text match."
    )
    
    parser.add_argument(
        "folder",
        type=str,
        help="Path to the folder containing Excel files"
    )
    
    parser.add_argument(
        "term",
        type=str,
        help="Search term (partial match allowed)"
    )

    args = parser.parse_args()

    matches = search_excel_with_pandas(args.folder, args.term)

    if matches:
        print("\n=== MATCHES FOUND ===")
        for m in matches:
            print(f"{m['file']} | {m['sheet']} | Row {m['row']} | Column {m['column']} -> {m['value']}")
    else:
        print("No matches found.")


if __name__ == "__main__":
    main()

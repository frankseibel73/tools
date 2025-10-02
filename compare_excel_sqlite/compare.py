import sqlite3
import pandas as pd
import yaml
import argparse

def load_config(config_path):
    with open(config_path, "r") as f:
        return yaml.safe_load(f)

def compare_dataframes(df_excel, df_sql, excel_file, sheet, table):
    results = []
    # Standardize column order and names
    df_excel.columns = df_excel.columns.astype(str).str.strip()
    df_sql.columns = df_sql.columns.astype(str).str.strip()
    
    # Compare column counts
    if len(df_excel.columns) != len(df_sql.columns):
        results.append(f"[{excel_file}:{sheet}] vs [{table}] - Column mismatch: "
                       f"Excel {len(df_excel.columns)} vs SQL {len(df_sql.columns)}")
    
    # Compare row counts
    if len(df_excel) != len(df_sql):
        results.append(f"[{excel_file}:{sheet}] vs [{table}] - Row mismatch: "
                       f"Excel {len(df_excel)} vs SQL {len(df_sql)}")
    
    # Align columns by name where possible
    common_cols = list(set(df_excel.columns) & set(df_sql.columns))
    df_excel_common = df_excel[common_cols].reset_index(drop=True)
    df_sql_common = df_sql[common_cols].reset_index(drop=True)

    # Compare cell-by-cell
    diffs = (df_excel_common != df_sql_common)
    for row, col in zip(*diffs.to_numpy().nonzero()):
        results.append(
            f"[{excel_file}:{sheet}] vs [{table}] "
            f"Mismatch at row {row+1}, column '{common_cols[col]}': "
            f"Excel='{df_excel_common.iloc[row, col]}' vs SQL='{df_sql_common.iloc[row, col]}'"
        )
    
    return results

def main(config_path, sqlite_db, output_path="comparison_report.txt"):
    config = load_config(config_path)
    conn = sqlite3.connect(sqlite_db)
    all_results = []

    for mapping in config["mappings"]:
        excel_file = mapping["excel_file"]
        for ws in mapping["worksheets"]:
            sheet = ws["name"]
            table = ws["table"]

            # Load data
            df_excel = pd.read_excel(excel_file, sheet_name=sheet)
            df_sql = pd.read_sql_query(f"SELECT * FROM {table}", conn)

            # Compare
            results = compare_dataframes(df_excel, df_sql, excel_file, sheet, table)
            if results:
                all_results.extend(results)
            else:
                all_results.append(f"[{excel_file}:{sheet}] vs [{table}] - ✅ All data matches")

    conn.close()

    # Write report
    with open(output_path, "w") as f:
        for line in all_results:
            f.write(line + "\n")
    print(f"Report saved to {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Compare Excel files to SQLite database for migration validation.")
    parser.add_argument("-c", "--config", default="config.yaml", help="Path to YAML config file (default: config.yaml)")
    parser.add_argument("-d", "--db", default="database.sqlite", help="Path to SQLite database file (default: database.sqlite)")
    parser.add_argument("-o", "--output", default="comparison_report.txt", help="Path to output report file (default: comparison_report.txt)")
    args = parser.parse_args()
    main(args.config, args.db, args.output)

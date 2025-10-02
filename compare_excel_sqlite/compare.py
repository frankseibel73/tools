#!/usr/bin/env python3
import argparse
import sqlite3
import pandas as pd
import os
import json
from tabulate import tabulate

def load_config(config_path):
    with open(config_path, "r") as f:
        return json.load(f)

def compare_excel_to_sqlite(db_path, excel_dir, config):
    conn = sqlite3.connect(db_path)
    results = []

    for table, mapping in config.items():
        file_name = mapping["file"]
        sheet_name = mapping["sheet"]

        excel_path = os.path.join(excel_dir, file_name)
        if not os.path.exists(excel_path):
            results.append([table, "MISSING FILE", file_name])
            continue

        # Load Excel
        df_excel = pd.read_excel(excel_path, sheet_name=sheet_name)
        # Load SQLite
        df_sql = pd.read_sql_query(f"SELECT * FROM {table}", conn)

        # Compare shapes
        match_rows = len(df_excel) == len(df_sql)
        match_cols = len(df_excel.columns) == len(df_sql.columns)

        # Compare data (row by row, col by col)
        match_data = df_excel.equals(df_sql)

        results.append([
            table,
            "OK" if match_rows else f"Row mismatch: Excel={len(df_excel)} DB={len(df_sql)}",
            "OK" if match_cols else f"Col mismatch: Excel={len(df_excel.columns)} DB={len(df_sql.columns)}",
            "OK" if match_data else "Data mismatch"
        ])

    conn.close()
    return results

def main():
    parser = argparse.ArgumentParser(description="Compare Excel files to SQLite DB")
    parser.add_argument("--db", required=True, help="Path to SQLite database")
    parser.add_argument("--excel-dir", required=True, help="Path to directory containing Excel files")
    parser.add_argument("--config", required=True, help="Path to JSON config mapping tables to Excel files/sheets")
    args = parser.parse_args()

    config = load_config(args.config)
    results = compare_excel_to_sqlite(args.db, args.excel_dir, config)

    print(tabulate(results, headers=["Table", "Rows", "Columns", "Data"], tablefmt="grid"))

if __name__ == "__main__":
    main()

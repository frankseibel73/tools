import sqlite3
import pandas as pd
import yaml
import argparse
from datetime import datetime

def load_config(config_path):
    with open(config_path, "r") as f:
        return yaml.safe_load(f)

def compare_dataframes(df_excel, df_sql, excel_file, sheet, table, ignore_missing_excel_cols=False):
    import re
    def clean_value(val):
        if pd.isna(val) or val is None:
            return ''
        s = str(val)
        # Remove Excel XML control codes and common control characters
        s = re.sub(r'_x000D_', '', s)
        s = s.replace('\r', '').replace('\n', '').replace('\t', '')
        # Remove other non-printable/control characters
        s = re.sub(r'[\x00-\x1F\x7F]', '', s)
        return s.strip()
    results = []
    # Standardize column order and names
    df_excel.columns = df_excel.columns.astype(str).str.strip()
    df_sql.columns = df_sql.columns.astype(str).str.strip()
    
    # Compare column counts
    if len(df_excel.columns) != len(df_sql.columns):
        excel_cols = set(df_excel.columns)
        sql_cols = set(df_sql.columns)
        missing_in_excel = sorted(list(sql_cols - excel_cols))
        missing_in_sql = sorted(list(excel_cols - sql_cols))
        msg = f"[{excel_file}:{sheet}] vs [{table}] - Column mismatch: " \
              f"Excel {len(df_excel.columns)} vs SQL {len(df_sql.columns)}"
        if missing_in_excel and not ignore_missing_excel_cols:
            msg += f"; Missing in Excel: {missing_in_excel}"
        if missing_in_sql:
            msg += f"; Missing in SQL: {missing_in_sql}"
        # Only append if not ignoring all missing columns in Excel
        if not (ignore_missing_excel_cols and missing_in_excel and not missing_in_sql):
            results.append(msg)
    
    # Compare row counts
    if len(df_excel) != len(df_sql):
        results.append(f"[{excel_file}:{sheet}] vs [{table}] - Row mismatch: "
                       f"Excel {len(df_excel)} vs SQL {len(df_sql)}")
    
    # Align columns by name where possible
    common_cols = list(set(df_excel.columns) & set(df_sql.columns))
    df_excel_common = df_excel[common_cols].reset_index(drop=True)
    df_sql_common = df_sql[common_cols].reset_index(drop=True)

    # Normalize blanks: treat NaN, None, and empty string as equivalent
    def normalize_blanks(df):
        return df.where(pd.notnull(df), '').replace({None: '', pd.NA: '', 'nan': ''})
    df_excel_common = normalize_blanks(df_excel_common)
    df_sql_common = normalize_blanks(df_sql_common)

    # Only compare cell-by-cell if shapes match
    if df_excel_common.shape == df_sql_common.shape:
        # Compare cell-by-cell, treating numerics and dates as equal if normalized values match
        for row in range(df_excel_common.shape[0]):
            for col in range(df_excel_common.shape[1]):
                val_excel = clean_value(df_excel_common.iloc[row, col])
                val_sql = clean_value(df_sql_common.iloc[row, col])
                # Treat both blanks as equal
                if (str(val_excel).strip() == '' and str(val_sql).strip() == '') or \
                   (pd.isna(val_excel) and pd.isna(val_sql)) or \
                   (val_excel is None and val_sql is None):
                    continue
                # Treat blank Excel as zero if SQL is zero
                if (str(val_excel).strip() == '' and str(val_sql).strip() in ['0', '0.0']) or \
                   (str(val_sql).strip() == '' and str(val_excel).strip() in ['0', '0.0']):
                    continue
                # Treat Excel NaT or blank as equal to SQL '0001-01-01' (default date)
                if (str(val_excel).strip() in ['', 'NaT'] and str(val_sql).strip() in ['','0001-01-01']) or \
                   (str(val_sql).strip() in ['', 'NaT'] and str(val_excel).strip() == ['','0001-01-01']):
                    continue
                # Treat Excel True as equal to TRUE
                if (str(val_excel).strip().lower() == 'true' and str(val_sql).strip().upper() == 'TRUE') or \
                   (str(val_sql).strip().lower() == 'true' and str(val_excel).strip().upper() == 'TRUE') or \
                   (str(val_excel).strip().lower() == 'false' and str(val_sql).strip().upper() == 'FALSE') or \
                   (str(val_sql).strip().lower() == 'false' and str(val_excel).strip().upper() == 'FALSE') or \
                   (str(val_excel).strip().lower() == 'true' and str(val_sql).strip().upper() in ['1',1]) or \
                   (str(val_sql).strip().lower() == 'true' and str(val_excel).strip().upper() in ['1',1]) or \
                   (str(val_excel).strip().lower() == 'false' and str(val_sql).strip().upper() in ['0',0]) or \
                   (str(val_sql).strip().lower() == 'false' and str(val_excel).strip().upper() in ['0',0]):
                    continue
                # Try to compare as dates
                try:
                    date_excel = pd.to_datetime(val_excel, errors='raise')
                    date_sql = pd.to_datetime(val_sql, errors='raise')
                    if date_excel.date() == date_sql.date():
                        continue
                except Exception:
                    pass
                # Compare as strings first
                if str(val_excel) == str(val_sql):
                    continue
                # If string comparison fails, try numeric
                try:
                    num_excel = float(val_excel) if str(val_excel).strip() != '' else 0.0
                    num_sql = float(val_sql) if str(val_sql).strip() != '' else 0.0
                    if pd.isna(num_excel) and pd.isna(num_sql):
                        continue
                    if num_excel == num_sql:
                        continue
                except (ValueError, TypeError):
                    pass
                results.append(
                    f"[{excel_file}:{sheet}] vs [{table}] "
                    f"Mismatch at row {row+1}, column '{common_cols[col]}': "
                    f"Excel='{val_excel}' vs SQL='{val_sql}'"
                )
    else:
        results.append(f"[{excel_file}:{sheet}] vs [{table}] - Skipped cell-by-cell comparison due to shape mismatch: "
                       f"Excel shape {df_excel_common.shape} vs SQL shape {df_sql_common.shape}")
    return results

def main(config_path, sqlite_db, output_path="comparison_report.txt", ignore_missing_excel_cols=False):
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
            results = compare_dataframes(df_excel, df_sql, excel_file, sheet, table, ignore_missing_excel_cols=ignore_missing_excel_cols)
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
    parser.add_argument("-o", "--output", default="comparison_report.txt", help="Base name for output report file (default: comparison_report.txt)")
    parser.add_argument("--ignore-missing-excel-cols", action="store_true", help="Ignore missing columns in Excel that are present in the database table.")
    args = parser.parse_args()
    # Add date and time stamp to output file name
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if "." in args.output:
        base, ext = args.output.rsplit('.', 1)
        output_file = f"{base}_{timestamp}.{ext}"
    else:
        output_file = f"{args.output}_{timestamp}"
    main(args.config, args.db, output_file, ignore_missing_excel_cols=args.ignore_missing_excel_cols)

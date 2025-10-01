import os
import openpyxl
import xlrd

def list_excel_sheets(folder_path):
    workbook_info = []

    for file in os.listdir(folder_path):
        if file.endswith(".xlsx") or file.endswith(".xls"):
            file_path = os.path.join(folder_path, file)
            try:
                if file.endswith(".xlsx"):
                    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                    sheets = wb.sheetnames
                else:  # .xls
                    wb = xlrd.open_workbook(file_path)
                    sheets = wb.sheet_names()

                for sheet in sheets:
                    workbook_info.append((file, sheet))

            except Exception as e:
                print(f"❌ Error reading {file}: {e}")

    return workbook_info

if __name__ == "__main__":
    folder = "path/to/your/folder"  # <-- change this
    results = list_excel_sheets(folder)

    print("Found the following workbooks and worksheets:")
    for file, sheet in results:
        print(f"📘 {file} -> 📝 {sheet}")

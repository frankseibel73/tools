# Excel-SQLite Tools

This repository contains a set of Python utilities to work with **Excel spreadsheets** and **SQLite databases**, packaged with Docker for isolated, reproducible execution.

## 📂 Repository Structure

Each script is stored in its own folder:

```
.
├── compare_excel_sqlite/      # Compare Excel vs SQLite contents
│   └── compare.py
├── list_excel_sheets/         # List all Excel workbooks and worksheets
│   └── list_sheets.py
├── Dockerfile                 # Build environment for running scripts
├── requirements.txt           # Python dependencies
└── README.md
```

## 🚀 Features

- Compare Excel spreadsheets against imported SQLite tables  
- Generate human-readable reports of mismatches (rows, columns, values)  
- List all workbooks and worksheets in a folder  
- Run inside Docker with a mapped folder for input/output  

## 🐳 Running with Docker

### 1. Build the image
```bash
docker build -t excel-sqlite-tools .
```

### 2. Run the container with a mapped folder

Assume your Excel/SQLite files live in `~/data` on your host:

```bash
docker run --rm -it -v ~/data:/data excel-sqlite-tools
```

Inside the container:
- `/app/repo/` → contains the scripts from this repo  
- `/data/` → is your mapped host folder with Excel/SQLite files  

---

## 📜 Scripts

### 🔎 1. List Excel Workbooks & Worksheets
Lists all Excel files and sheet names in a specified folder.

```bash
python /app/repo/list_excel_sheets/list_sheets.py /data
```

Example output:
```
📘 Customers.xlsx -> 📝 Customers
📘 Customers.xlsx -> 📝 CustomerLocations
📘 Orders.xls -> 📝 OrderHeader
📘 Orders.xls -> 📝 OrderItems
```

---

### 📊 2. Compare Excel Data vs SQLite
Compares data from Excel sheets with tables in a SQLite database, based on a YAML config.

```bash
python /app/repo/compare_excel_sqlite/compare.py /data/config.yaml /data/database.sqlite
```

**Example `config.yaml`:**
```yaml
mappings:
  - excel_file: Customers.xlsx
    worksheets:
      - name: Customers
        table: Customers
      - name: CustomerLocations
        table: CustomerLocations
  - excel_file: Orders.xlsx
    worksheets:
      - name: OrderHeader
        table: OrderHeader
      - name: OrderItems
        table: OrderItems
```

**Example Output (`comparison_report.txt`):**
```
[Customers.xlsx:Customers] vs [Customers] - ✅ All data matches
[Customers.xlsx:CustomerLocations] vs [CustomerLocations] - Row mismatch: Excel 200 vs SQL 198
```

---

## ⚙️ Dependencies

All dependencies are installed via `requirements.txt`:

```
pandas
openpyxl
xlrd
pyyaml
```

SQLite support is built into Python.

---

## 📌 Notes

- To update the scripts, rebuild the Docker image after pulling changes.  
- You can extend the `Dockerfile` to automatically run a specific script using `ENTRYPOINT`.  
- All input/output files should be placed in your mapped folder (e.g., `~/data`).  

# ATS BOM Automation Tool v1.0

<p align="center">
  <img src="ats_logo.png" alt="ATS Group Logo" width="280"/>
</p>

<p align="center">
  <strong>Convert Vault CSV Exports into Clarity ERP Import Files</strong><br>
  <em>Developed by ATS Automation Team</em>
</p>

---

# ATS BOM Automation Tool

## Overview

ATS BOM Automation Tool converts Autodesk Vault BOM CSV exports into Clarity ERP import files.

The application generates:

1. Item Import Excel File
2. BOM Import Excel File

Workflow:

Vault CSV → ATS BOM Automation Tool → Clarity ERP Excel Files

---

## Features

* Select Vault CSV file
* Validate input data
* Generate Item Import file
* Generate BOM Import file
* Export Excel files ready for Clarity ERP
* Simple desktop UI

---

## Input

Source File:

```text
Vault_BOM.csv
```

Generated Files:

```text
Vault_BOM-ItemImport.xlsx
Vault_BOM_BOM_Import_Clarity.xlsx
```

---

## Project Structure

```text
project/
│
├── ATS_BOM_Automation_Tool.py
├── ats_logo.png
├── README.md
└── requirements.txt
```

Note:

All conversion logic is integrated into the main application.

No external Python scripts are required.

---

## Installation (Developer)

Install Python dependencies:

```bash
pip install pandas openpyxl xlsxwriter customtkinter pillow
```

Run application:

```bash
py ATS_BOM_Automation_Tool.py
```

---

## Build EXE

Install PyInstaller:

```bash
pip install pyinstaller
```

Generate executable:

```bash
pyinstaller --onefile --windowed --name ATS_BOM_Automation_Tool ATS_BOM_Automation_Tool.py
```

Output:

```text
dist/
└── ATS_BOM_Automation_Tool.exe
```

---

## User Installation

No Python installation required.

User only needs:

```text
ATS_BOM_Automation_Tool.exe
```

Double-click and run.

---

## Process Flow

```text
1. Export BOM from Autodesk Vault

2. Open ATS BOM Automation Tool

3. Select CSV File

4. Click Generate

5. Application creates:
   - Item Import Excel
   - BOM Import Excel

6. Import generated files into Clarity ERP
```

---

## Troubleshooting

### Python not found

```bash
py --version
```

### Missing library

```bash
pip install -r requirements.txt
```

### EXE not opening

Right click:

```text
More Info
→ Run Anyway
```

---

## Version

Version: 1.0

Developed for ATS Engineering Team.

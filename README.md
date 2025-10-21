# 🧾 SharePoint Weekly Report Generator

This Python project automatically downloads Excel workbooks from **Microsoft SharePoint**, processes the data, and generates a **weekly PDF report**.  
The report includes simulator utilization, discrepancy summaries, preventive maintenance, and RTMS session data.

---

## 🚀 Features

- 🔗 **Download Excel files from SharePoint** using Microsoft Graph API (with an authentication with Bearer token)
- 📊 **Parse and merge data** from multiple sources:
  - SH Duty Logbook
  - Discrepancy and Preventive Maintenance logs
  - RTMS Logbook
- 🧮 **Generate PDF reports** with tables, charts, and legends
- 🧹 Automatically **delete temporary Excel files** after processing

---

### 📁 Project Folder Structure

```plaintext
📁 sharepoint-report-generator/
│
├── 📂 src/                       # Source code
│   ├── main.py                   # Main script that orchestrates everything
│   ├── dataParser.py             # Handles SharePoint download + data parsing
│   ├── report.py                 # Builds the report (tables, charts, etc.)
│   └── __init__.py               # (optional) to mark the folder as a Python package
│
├── 📂 config/                    # Configuration files
│   └── env.env                   # Environment variables (SharePoint, MSAL, paths)
│
├── 📂 temp/                      # Temporary files (Excel downloads, etc.)
│   └── (auto-created/deleted by script)
│
├── 📂 output/                    # Generated reports
│   ├── Weekly_Report_21_2025.pdf
│   └── ...
│
├── 📂 assets/                    # Icons, screenshots, logos, etc.
│   ├── ajt_official.ico
│   └── preview.png
│
├── .gitignore                    # Ignore temp files, env, caches, etc.
├── requirements.txt              # Python dependencies
├── README.md                     # Project documentation
└── LICENSE                       # License (MIT or other)

---
###🖥️ Run the Script

Run the generator directly:

python main.py


Or specify week and year with a kinter GUI:

python main.py --year 2025 --week 21

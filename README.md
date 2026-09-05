# ⚡ Automated Excel Report Generator (Python & OpenPyXL)

An automated data processing and reporting tool built with **Python**, **Pandas**, and **OpenPyXL**. It converts unorganized, raw data (CSV/JSON) into fully formatted, presentation-ready Excel reports—reducing manual report generation time from hours to seconds.

---

## 🎯 Business Problem

Operations and finance teams often waste hours manually copying, pasting, cleaning, and formatting daily/weekly report files. Manual manipulation introduces high human error rates, inconsistent formatting, date/encoding mismatches, and broken formulas.

This script automates the entire end-to-end reporting pipeline: data extraction, cleaning, date normalization, cell formatting, dynamic formula insertion, and visual styling.

---

## ✨ Key Features & Capabilities

* **Automated Data Ingestion:** Safely handles varied raw file encodings (`UTF-8`, `Latin-1`, `cp1252`) without script failure.
* **Data Cleansing & Validation:** Strips whitespace, normalizes date strings using `dateutil`, applies regex text cleaning, and handles missing values.
* **Automated Excel Formatting:**
  * Auto-adjusts column widths based on cell content length.
  * Formats currency, percentage, and date values natively in Excel (`$#,##0.00`, `YYYY-MM-DD`).
  * Applies professional headers, border styles, and conditional highlighting using `openpyxl`.
* **Formula Injection:** Automatically inserts dynamic summary formulas (`SUM`, `AVERAGE`, `COUNTIF`) at the footer of report tables.
* **Robust Error Handling & Logging:** Implements structured logging (`logging` module) and path handling (`pathlib`) to track execution and audit pipeline failures.

---

## 🛠️ Technical Stack

* **Language:** Python 3.10+
* **Data Manipulation:** Pandas, NumPy
* **Excel Engine:** OpenPyXL, Dateutil, Pathlib, Regex
* **Core Concepts:** Process Automation, Data Quality, File I/O, Error Handling, Audit Logging

---

## 🚀 Getting Started

```bash
# Clone the repository
git clone [https://github.com/davidstocco2024-cell/Excel-report-generator.git](https://github.com/davidstocco2024-cell/Excel-report-generator.git)
cd Excel-report-generator

# Install dependencies
pip install pandas openpyxl python-dateutil

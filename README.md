# Excel Report Generator for Construction Companies

**Automate expense report generation for construction and real estate projects.**

**Turn messy CSV/JSON data into professional Excel reports in seconds.**

---

## The Problem

Construction companies waste many hours every month manually cleaning expense data from suppliers and subcontractors. Inconsistent formats, wrong dates, and encoding issues make the process slow and error-prone.

## The Solution

This tool automates the complete process:
- Reads messy CSV and JSON files
- Automatically cleans data and fixes encoding issues
- Intelligently parses dates (multiple formats)
- Generates clean, professional Excel reports with formatting and charts

**Result:** What used to take 4-6 hours now takes less than 1 minute.

---

## Key Features

- Multi-format support (CSV, JSON, Excel)
- Smart encoding detection (UTF-8, Latin-1, CP1252)
- Intelligent date parsing with `dateutil`
- Advanced data cleaning using regex
- Professional Excel reports with styling, formulas, and charts
- Full logging and robust error handling
- Command-line interface with `click`

---

## Quick Start

```bash
git clone https://github.com/davidstocco2024-cell/Excel-report-generator.git
cd Excel-report-generator
pip install -r requirements.txt
python main.py --input datos/gastos_abril.csv --output reportes/reporte_abril.xlsx



Technologies
Python 3.10+
openpyxl
python-dateutil
python-dotenv
click

Commercial Use
This repository is private for demonstration purposes only.
For commercial licensing, customization, or implementation services, please contact me.

Contact
David Stocco
Python Automation Specialist for Construction Companies
📧 dstoccoanalytics@gmail.com
📍 Mendoza, Argentina (Remote friendly - Australia & USA)

Made for construction companies in Mendoza and beyond.

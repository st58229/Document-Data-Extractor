# Big Data Processing App

A desktop application built with Python and Tkinter for processing and extracting structured data from **Excel**, **PDF**, and **Word** files. The tool supports conversion between CSV/JSON formats, selective attribute extraction, and merging of JSON files.

## Features

- ✅ Select individual files or entire folders for processing
- 📄 Extract data from specified Excel cells
- 📄 Extract values from Word or PDF files using keywords
- 🔄 Convert CSV ↔ JSON (configurable delimiter and decimal)
- 🔧 Define and import/export extraction settings
- 🧠 Merge multiple JSON files into one
- 📊 Track progress with a built-in progress bar
- 📁 View detailed file info (count, size, paths)
- 📂 Export processed data to JSON or convert it to CSV
- ℹ️ Easy-to-use graphical interface (Tkinter)

---

## Installation

### Prerequisites

Ensure you have Python 3.8+ installed. Then install the required libraries:

```bash
pip install pandas openpyxl PyMuPDF python-docx
```

---

## Usage

1. **Run the application**:

```bash
python app.py
```

2. **Main actions** (GUI-based):
   - Select files or folders for processing
   - Configure extraction settings for:
     - Excel (specific cells → attributes)
     - PDF/Word (keywords → attributes)
   - Export/import your settings to/from `.json`
   - Start processing and export results to JSON
   - Optionally convert resulting JSON into CSV

---

## JSON Structure

After processing, the exported JSON will be structured as:

```json
{
  "file_path_1": {
    "Attribute1": "Value1",
    "Attribute2": "Value2"
  },
  "file_path_2": {
    "Attribute1": "Value3",
    "Attribute2": "Value4"
  }
}
```

---

## Menu Overview

- **CSV -> JSON**
  - Convert structured CSV/Excel to JSON
  - Set custom delimiter or decimal separator
- **Merge JSON**
  - Merge multiple JSON files into one
- **JSON -> CSV**
  - Convert processed JSON into structured CSV
- **About**
  - Displays app version and author info

---

## Author

**Tomáš Fráňa**  
`st58229@upce.cz`  
Version: `1.0`

---

## License

This project is distributed for educational and demonstration purposes.  
Feel free to modify and adapt.

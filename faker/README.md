# Invoice Generator Script

A Python-based tool for generating synthetic invoices in various formats (Excel, Word, PDF) using template files and Czech business data via the Faker library.

## Features

- 📄 Generate random invoices for testing/demo purposes
- 📁 Supports Excel templates (`.xlsx`) with customizable fields
- 📄 Generate Word (`.docx`) invoices using placeholders
- 🧾 Create stylized PDF invoices with UTF-8 font support
- 🧠 Uses Czech locale (`cs_CZ`) for realistic fake company info
- 🕒 Each batch is uniquely timestamped

---

## Installation

### Prerequisites

Install Python 3.8+ and the required libraries:

```bash
pip install faker openpyxl python-docx fpdf
```

---

## Usage

Run the script using the terminal:

```bash
python invoice_generator.py
```

You will be prompted to enter:
- Invoice type: `excel`, `word`, or `pdf`
- For Excel: Choose a template number (1, 2, or 3)
- Number of invoices to generate

The generated files will be saved to the `generated_invoices` folder by default.

---

## Template Requirements

Ensure the following templates are present in the same folder as the script:
- `Sablona_Excel_Faktura1.xlsx`
- `Sablona_Excel_Faktura2.xlsx`
- `Sablona_Excel_Faktura3.xlsx`
- `Sablona_Word_Faktura1.docx`

Placeholders in Word templates must follow this format:

```
{INVOICE_NUMBER}, {DATE_ISSUED}, {DUE_DATE}, {COMPANY_NAME}, {ADDRESS},
{TOTAL_AMOUNT}, {MONTH_NAME}, {ICO}, {DIC}, {PHONE_NUMBER},
{BANK_NAME}, {BANK_ACCOUNT}, {VS}, {CITY_NAME}
```

---

## Output Files

The generated invoices are saved in the `generated_invoices/` directory as:
- `invoice_<timestamp>_<index>.xlsx`
- `invoice_<timestamp>_<index>.docx`
- `invoice_<timestamp>_<index>.pdf`

---

## Author

**Tomáš Fráňa**  
`st58229@upce.cz`  
Version: `1.0`

---

## License

This project is intended for testing and educational purposes.

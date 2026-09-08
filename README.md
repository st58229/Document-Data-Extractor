# Document-Data-Extractor
An open-source, GUI-based Python application for extracting data from heterogeneous business documents (Excel, Word, PDF) into structured JSON/CSV formats in resource-constrained environments.

This tool is specifically tailored for small and medium-sized enterprises (SMEs) and non-technical users, providing a zero-code workflow to transform semi-structured office documents (Excel, PDF, Word) into structured JSON or CSV formats ready for downstream analysis or database insertion.

## Key Features

* **Zero-Code GUI Extraction Mapping:** Define data extraction rules visually through the interface without requiring database administration or programming skills.


* **Format Support:** Native processing capabilities for `.xlsx`, `.docx`, `.pdf`, and `.csv` files.


* **Sequential Batch Processing:** Operates with $O(N)$ time complexity, ensuring predictable and bounded memory usage even when processing batches of up to 10,000 documents on standard office hardware.


* **Error Recovery and Auditing:** Generates detailed error logs for missing keywords or corrupted files, allowing users to correct mapping errors for specific files without terminating the entire batch process.


* **Targeted Data Reduction:** Extracts only user-defined fields (e.g., Invoice Number, Total, Date), discarding XML and layout overhead, which results in significantly smaller structured output files.



## Repository Structure

* `src/` – The complete source code of the Python application.


* `example_configs/` – Example JSON configuration files demonstrating extraction rules for Excel and text-based documents.


* `synthetic_data_generator/` – A Python script utilizing the *Faker* library to generate synthetic document datasets. This ensures full reproducibility of the experimental results and performance metrics.



## Requirements

* Python 3.10+


* `pandas`

* `PyMuPDF`

* `openpyxl`

* `python-docx`

* `Tkinter` (Standard Python library for the GUI)



## Installation and Usage

1. Clone this repository:
```bash
git clone https://github.com/your-username/Document-Data-Extractor.git

```


2. Install the required dependencies:
```bash
pip install -r requirements.txt

```


3. Launch the application:
```bash
python main.py

```


4. **Using the GUI:** Select your target directory, choose the file format (Excel or Word/PDF), define your extraction mapping rules (e.g., cell coordinates or keyword pairs), and execute the batch process to export directly to JSON or CSV.



## Reproducibility and Synthetic Data

To replicate the performance and memory footprint experiments discussed in our research, navigate to the `synthetic_data_generator/` directory. Run the generation script to create a localized synthetic dataset of business documents (e.g., invoices) in `.xlsx`, `.pdf`, and `.docx` formats.

## Academic Citation

If you utilize this framework in your research, please cite our paper:

> Frana, T., Lnenicka, M., & Horak, O. (2026). A Python-based application for document data extraction and preparation in resource-constrained environments. *Array*.
> 
>

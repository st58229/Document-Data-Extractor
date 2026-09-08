import os
import random
from faker import Faker
from openpyxl import load_workbook
from fpdf import FPDF
from docx import Document
from datetime import datetime


def generate_invoices(num_invoices, invoice_type, output_folder="generated_invoices"):
    # Initialize Faker for Czech data
    fake = Faker("cs_CZ")
    os.makedirs(output_folder, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # Unique timestamp for the invoice batch

    # Templates for generating invoices
    excel_template1 = "Sablona_Excel_Faktura1.xlsx"
    excel_template2 = "Sablona_Excel_Faktura2.xlsx"
    excel_template3 = "Sablona_Excel_Faktura3.xlsx"
    word_template = "Sablona_Word_Faktura1.docx"

    # Font path fix for PDF generation
    font_path = "C:\\Windows\\Fonts\\arial.ttf"  # Path to TTF file on Windows
    if not os.path.exists(font_path):
        font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"  # Alternative for Linux/Mac

    # Actual invoice generation
    for i in range(num_invoices):
        invoice_number = random.randint(100000, 999999)
        date_issued = datetime.today().strftime("%d.%m.%Y")
        due_date = (datetime.today().replace(day=28)).strftime("%d.%m.%Y")
        company_name = fake.company()
        address = fake.address().replace("\n", ", ")
        total_amount = round(random.uniform(500, 50000), 2)
        ico = fake.ssn()
        bank_list = [
            "Česká spořitelna", "Komerční banka", "ČSOB", "MONETA Money Bank",
            "Fio banka", "Air Bank", "UniCredit Bank", "Raiffeisenbank"
        ]
        bank_name = random.choice(bank_list)
        bank_account = f"{random.randint(1000000000, 9999999999)}/{random.randint(1000, 9999)}"  # Bank account number in format XXXXXXXX/YYYY
        variable_symbol = random.randint(1000000000, 9999999999)
        city_name = fake.city()
        phone_number = fake.phone_number()
        months_cs = {
            "January": "Leden", "February": "Únor", "March": "Březen",
            "April": "Duben", "May": "Květen", "June": "Červen",
            "July": "Červenec", "August": "Srpen", "September": "Září",
            "October": "Říjen", "November": "Listopad", "December": "Prosinec"
        }
        month_name = months_cs[datetime.now().strftime("%B")]

        if invoice_type == "excel1":
            # Modify Excel template
            wb = load_workbook(excel_template1)
            sheet = wb.active

            # Configurable values
            sheet["E1"] = invoice_number

            # Supplier
            sheet["E4"] = company_name
            sheet["E5"] = address
            sheet["E6"] = ico
            sheet["E7"] = ""
            sheet["E8"] = f"CZ{ico}"
            sheet["E9"] = bank_name
            sheet["E10"] = bank_account
            sheet["E11"] = ""
            sheet["E12"] = phone_number

            sheet["E15"] = variable_symbol
            sheet["E16"] = due_date
            sheet["E17"] = total_amount

            sheet["A20"] = month_name

            sheet["A25"] = city_name
            sheet["A26"] = date_issued

            # Save Excel invoice
            excel_file = os.path.join(output_folder, f"invoice_{timestamp}_{i + 1}.xlsx")
            wb.save(excel_file)
            print(f"Generovaná faktura uložena jako: {excel_file}")

        elif invoice_type == "excel2" or invoice_type == "excel3":

            # Modify Excel template
            if invoice_type == "excel2":
                wb = load_workbook(excel_template2)
            else:
                wb = load_workbook(excel_template3)
            sheet = wb.active

            sheet["C1"] = invoice_number

            # Supplier
            sheet["E4"] = company_name
            sheet["E5"] = address
            sheet["E6"] = ico
            sheet["E7"] = f"CZ{ico}"
            sheet["E9"] = bank_name
            sheet["E10"] = bank_account
            sheet["E11"] = ""
            sheet["E13"] = phone_number

            sheet["B15"] = variable_symbol
            sheet["B16"] = due_date
            sheet["B17"] = total_amount

            sheet["F16"] = month_name

            sheet["B21"] = city_name
            sheet["B22"] = date_issued

            # Save Excel invoice
            excel_file = os.path.join(output_folder, f"invoice_{timestamp}_{i + 1}.xlsx")
            wb.save(excel_file)
            print(f"Generovaná faktura uložena jako: {excel_file}")

        elif invoice_type == "word":
            # Modify Word template
            doc = Document(word_template)
            for paragraph in doc.paragraphs:
                for run in paragraph.runs:  # Preserve formatting and images
                    run.text = run.text.replace("{INVOICE_NUMBER}", str(invoice_number))
                    run.text = run.text.replace("{DATE_ISSUED}", date_issued)
                    run.text = run.text.replace("{DUE_DATE}", due_date)
                    run.text = run.text.replace("{COMPANY_NAME}", company_name)
                    run.text = run.text.replace("{ADDRESS}", address)
                    run.text = run.text.replace("{TOTAL_AMOUNT}", f"{total_amount} Kč")
                    run.text = run.text.replace("{MONTH_NAME}", month_name)
                    run.text = run.text.replace("{ICO}", ico)
                    run.text = run.text.replace("{DIC}", f"CZ{ico}")
                    run.text = run.text.replace("{PHONE_NUMBER}", phone_number)
                    run.text = run.text.replace("{BANK_NAME}", bank_name)
                    run.text = run.text.replace("{BANK_ACCOUNT}", bank_account)
                    run.text = run.text.replace("{VS}", f"{variable_symbol}")
                    run.text = run.text.replace("{CITY_NAME}", city_name)

            word_file = os.path.join(output_folder, f"invoice_{timestamp}_{i + 1}.docx")
            doc.save(word_file)
            print(f"Generovaná faktura uložena jako: {word_file}")

        elif invoice_type == "pdf":
            # Create PDF invoice with UTF-8 support
            pdf = FPDF()
            pdf.add_page()
            pdf.add_font('CustomFont', '', font_path, uni=True)
            pdf.set_font("CustomFont", '', 16)
            pdf.cell(200, 10, txt="Faktura", ln=True, align='C')
            pdf.set_font("CustomFont", '', 12)
            pdf.cell(200, 10, txt=f"Faktura č. {invoice_number}", ln=True)
            pdf.cell(200, 10, txt=f"Datum vystavení: {date_issued}", ln=True)
            pdf.cell(200, 10, txt=f"Datum splatnosti: {due_date}", ln=True)
            pdf.cell(200, 10, txt=f"Dodavatel: {company_name}", ln=True)
            pdf.cell(200, 10, txt=f"IČO: {ico}", ln=True)
            pdf.cell(200, 10, txt=f"Adresa: {address}", ln=True)
            pdf.cell(200, 10, txt=f"Bankovní účet: {bank_account} ({bank_name})", ln=True)
            pdf.cell(200, 10, txt=f"Variabilní symbol: {variable_symbol}", ln=True)
            pdf.cell(200, 10, txt=f"Telefon: {phone_number}", ln=True)
            pdf.cell(200, 10, txt=f"Celková částka: {total_amount} Kč", ln=True)
            pdf_file = os.path.join(output_folder, f"invoice_{timestamp}_{i + 1}.pdf")
            pdf.output(pdf_file)
            print(f"Generovaná faktura uložena jako: {pdf_file}")

        else:
            print("Neplatný typ faktury.")
            break


if __name__ == "__main__":
    invoice_type = input("Enter invoice type (excel/word/pdf): ").strip().lower()
    if invoice_type == "excel":
        invoice_type = "excel" + input("Enter template type (1, 2, 3): ").strip().lower()
    num_invoices = int(input("How many invoices would you like to generate? "))

    generate_invoices(num_invoices, invoice_type)

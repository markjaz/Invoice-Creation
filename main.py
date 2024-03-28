"""
Read data from various Excel files and create a PDF invoice presenting the
data.
"""
from fpdf import FPDF
import glob
from pathlib import Path
import pandas as pd

# Create a Python list containing the file names of items in '/invoices'.
filepaths = glob.glob('Excel files/*.xlsx')
print(filepaths)
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    # Extract the filename without extension.
    filename = Path(filepath).stem
    invoice_number, invoice_date = filename.split("-")

    pdf.set_font("Helvetica", size=16, style="B")
    # The 'ln=1' parameter moves the cursor to the next line below the cell.
    pdf.cell(w=50, h=8, txt=f"Invoice number: ", align='R', border=True,
             ln=0)
    pdf.cell(w=50, h=8, txt=invoice_number, align='L', border=True, ln=1)

    pdf.set_font("Helvetica", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice date: ", align='R', border=True, ln=0)
    pdf.cell(w=50, h=8, txt=invoice_date, align='L', border=True, ln=1)

    pdf.output(f"PDF Invoices/{filename}.pdf")
    print(f"Created PDF Invoices/{filename}.pdf")

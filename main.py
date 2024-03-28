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
    invoice_number = filename.split("-")[0]
    pdf.set_font("Helvetica", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no: {invoice_number}")
    pdf.output(f"PDF Invoices/{filename}.pdf")
    print(f"Created PDF Invoices/{filename}.pdf")

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
# print(filepaths)
for filepath in filepaths:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    # Extract the filename without extension.
    filename = Path(filepath).stem

    # Get invoice number and date from the filename.
    invoice_number, invoice_date = filename.split("-")
    # Create a sort of header containing invoice number and date.
    pdf.set_font("Helvetica", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice number: ", align='R', border=False)
    pdf.cell(w=50, h=8, txt=invoice_number, align='L', border=False, ln=1)
    # The 'ln=1' parameter moves the cursor to the next line below the cell.
    pdf.set_font("Helvetica", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice date: ", align='R', border=False)
    pdf.cell(w=50, h=8, txt=invoice_date, align='L', border=False)

    # Put a logo image on the page
    pdf.image('1115-_5D1-9629-fu jin.png', x=155, y=2, w=50)

    # Note that if we were OK with the preset column names and just needed
    # to replace the underscores with spaces and capitalize the names,
    # a list comprehension would do it.  Since 'df.columns' is an iterable
    # object, we can replace underscores and capitalize with this list
    # comprehension:
    # col_names = [item.replace("_", " ").title() for item in df.columns]

    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    # Create the column header row
    pdf.set_xy(10, 40)
    pdf.set_font("Helvetica", size=12, style='B')
    pdf.set_text_color(64, 64, 64)
    pdf.cell(w=25, h=8, txt='Product ID', align='L', border=True)
    pdf.cell(w=85, h=8, txt='Product Name/Description', align='L', border=True)
    pdf.cell(w=20, h=8, txt='Quantity', align='R', border=True)
    pdf.cell(w=30, h=8, txt='Unit Price', align='R', border=True)
    pdf.cell(w=30, h=8, txt='Total Price', align='R', ln=1, border=True)
    # Fill in the actual data for each column
    total_amt_due = 0
    for index, row in df.iterrows():
        pdf.set_font("Helvetica", size=10)
        pdf.set_text_color(96, 96, 96)
        pdf.cell(w=25, h=8, txt=str(row['product_id']), align='L',
                 border=True)
        pdf.cell(w=85, h=8, txt=str(row['product_name']), align='L',
                 border=True)
        pdf.cell(w=20, h=8, txt=str(row['amount_purchased']), align='R',
                 border=True)
        pdf.cell(w=30, h=8, txt=str(f'$ {row['price_per_unit']:,.2f}'),
                 align='R',
                 border=True)
        pdf.cell(w=30, h=8, txt=str(f'$ {row['total_price']:,.2f}'), align='R',
                 ln=1, border=True)
        total_amt_due += row['total_price']

    # Calculate the sum due for all items.
    pdf.set_font("Helvetica", size=10, style='B')
    pdf.set_text_color(96, 96, 96)
    pdf.cell(w=160, h=8, txt='Total Amount Due:', align='R', border='True')
    pdf.cell(w=30, h=8,
             txt=str(f'$ {total_amt_due:,.2f}'), align='R', border='True')
    pdf.output(f"PDF Invoices/{filename}.pdf")
    print(f"Created PDF Invoices/{filename}.pdf")

import glob
from pathlib import Path

import pandas as pd
from fpdf import FPDF

# Get all filepaths of Excel files in the invoices folder
filepaths = glob.glob("invoices/*.xlsx")

# Iterate over each Excel file
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    # Create PDF instance
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    # Extract the filename and invoice number
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    
    # Set font and write invoice number to the PDF
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}")
    
    # Output the PDF to the specified directory
    pdf.output(f"PDFs/{filename}.pdf")

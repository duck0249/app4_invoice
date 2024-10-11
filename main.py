import glob
from pathlib import Path

import pandas as pd
from fpdf import FPDF

# Get all filepaths of Excel files in the invoices folder
filepaths = glob.glob("invoices/*.xlsx")

# Iterate over each Excel file
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1", engine="openpyxl")
    
    # Create PDF instance
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    # Extract the filename, invoice number and date
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    
    # Set font and write invoice number, date to the PDF
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}",ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {date}",ln=1)
    pdf.ln()

    # Create table and set up.
    pdf.set_font(family="Times", size=12)
    col_widths = [30, 50, 40, 40, 30]

    for col in df.columns:
    	pdf.cell(col_widths[df.columns.get_loc(col)],10, col.replace("_", " ").title(), 1, 0, "L")
    pdf.ln()

    for index, row in df.iterrows():
    	for i, item in enumerate(row):
    		pdf.cell(col_widths[i], 10, str(item).replace("_"," ").title(), 1, 0, "L")
    	pdf.ln()



    # Output the PDF to the specified directory
    pdf.output(f"PDFs/{filename}.pdf")

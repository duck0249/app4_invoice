import glob
from pathlib import Path

import pandas as pd
from fpdf import FPDF

# Get all filepaths of Excel files in the invoices folder
filepaths = glob.glob("invoices/*.xlsx")

# Iterate over each Excel file
for filepath in filepaths:
    
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

    df = pd.read_excel(filepath, sheet_name="Sheet 1", engine="openpyxl")

    columns = [column.replace("_", " ").title() for column in df.columns]
    
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=65, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
    	pdf.set_font(family="Times", size=10)
    	pdf.set_text_color(80, 80, 80)
    	pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
    	pdf.cell(w=65, h=8, txt=str(row["product_name"]), border=1)
    	pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
    	pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
    	pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total_sum = df["total_price"].sum()

    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=65, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)
    pdf.ln()

    # Add total sum
    pdf.set_font(family="Times", size=18, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum} Euros.", ln=1)
    
    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"Pythonhow")
    pdf.image("pythonhow.png", w=10)

	# Output the PDF to the specified directory
    pdf.output(f"PDFs/{filename}.pdf")

import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Creating a python list with the filepaths as we have multiple files
filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    # Sheet name required as argument for Excel files
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Create a pdf for each file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    pdf.set_font(family="Times", style="B", size=22)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=50, h=10, txt=f"Invoice nr. {invoice_nr} ", align="L", ln=1)
    pdf.output(f"PDFs/{filename}.pdf")
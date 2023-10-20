import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Creating a python list with the filepaths as we have multiple files
filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:

    # Create a pdf for each file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Get the invoice number on top of pdf
    pdf.set_font(family="Times", style="B", size=22)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=50, h=10, txt=f"Invoice nr. {invoice_nr} ", align="L", ln=1)

    # Get the date for each pdf
    pdf.set_font(family="Times", style="B", size=20)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=50, h=10, txt=f"Date {date} ", align="L", ln=1)

    # Get the contents of each excel into a pdf file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header; an iterable index object
    columns = df.columns
    # Use list comprehension to capitalize column titles
    columns = [item.replace("_", " "). title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=row["product_name"], border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)


    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=8)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add total sum
    pdf.set_font(family="Times", size=16, style = "B")
    pdf.cell(w=30, h=8,txt=f"The total price is {total_sum}", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", size=12)
    pdf.cell(w=30, h=8, txt=f"Pythonhow", ln=1)
    pdf.image("pythonhow.png", w=5)

    pdf.output(f"PDFs/{filename}.pdf")
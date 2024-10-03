import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    # create empty pdfs
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # get filename for invoice number and date
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # set font, invoice number and date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    pdf.cell(w=30, h=8, ln=1)

    # read files
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add header
    columns = df.columns
    columns = [item.replace("_", " ") for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80,80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows to table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # get total sum
    total_sum = sum(df["total_price"])
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    pdf.cell(w=30, h=16, ln=1)

    # Final statement
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(0)
    pdf.cell(w=30, h=8, txt=f"The total price if {total_sum} Euros.", ln=1)

    # name and logo
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=35, h=8, txt="MadeWithPython")
    pdf.image("pythonhow.png", w=10)

    # generate pdfs
    pdf.output(f"pdfs/{filename}.pdf")
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for file in filepaths:
    pdf = FPDF(orientation="portrait", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    file_name: str = Path(file).stem

    no, date_str = file_name.split("-")
    pdf.cell(w=50, h=10, txt=f"Invoice #{no}", border=0, ln=1, align="L")
    pdf.cell(w=50, h=10, txt=f"Date: {date_str}", border=0, ln=2, align="L")

    df = pd.read_excel(file, sheet_name="Sheet 1")
    # Render headers
    headers = list(df.columns)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(50, 50, 50)
    for header in headers:
        pdf.cell(w=70 if header == "product_name" else 30, h=8,
                 txt=header.replace("_", " ").title(), border=1,
                 ln=1 if header == "total_price" else 0, align="L")
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10, style="")
        pdf.set_text_color(100, 100, 100)
        pdf.cell(w=30, h=8, txt=str(
            row["product_id"]), border=1, ln=0, align="L")
        pdf.cell(w=70, h=8, txt=row["product_name"],
                 border=1, ln=0, align="L")
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]),
                 border=1, ln=0, align="R")
        pdf.cell(w=30, h=8, txt=str(float(row["price_per_unit"])),
                 border=1, ln=0, align="R")
        pdf.cell(w=30, h=8, txt=str(float(
            row["total_price"])), border=1, ln=1, align="R")
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=160, h=8, border=1, ln=0, align="L")
    pdf.cell(w=30, h=8, txt=str(
        float(df["total_price"].sum())), border=1, ln=1, align="R")

    pdf.set_font(family="Times", size=12)
    pdf.cell(
        w=30, h=8, txt=f"The total price is: {float(df['total_price'].sum())}", ln=1)
    pdf.set_font(family="Times", size=14)
    pdf.cell(w=26, h=12, txt=f"PythonHow")
    pdf.image("images/pythonhow.png", w=12)

    pdf.output(f"pdfs/{file_name}.pdf")

import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for file in filepaths:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    pdf = FPDF(orientation="portrait", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    file_name: str = Path(file).stem

    no, date_str = file_name.split("-")
    pdf.cell(w=50, h=10, txt=f"Invoice #{no}", border=0, ln=1, align="L")
    pdf.cell(w=50, h=10, txt=f"Date: {date_str}", border=0, ln=0, align="L")
    pdf.output(f"pdfs/{file_name}.pdf")

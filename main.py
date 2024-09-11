from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path
filepaths = glob.glob("Excel_Sheets/*.xlsx")

for filepath in filepaths :
    df = pd.read_excel(filepath,sheet_name='Sheet 1')
    invoice = Path(filepath).stem.split('-')[0]
    date = Path(filepath).stem.split('-')[1]

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times', size=14, style='B')
    pdf.cell(w=0, h=14, txt=f'Invoice Number : {invoice}')
    pdf.cell(w=0, h=14, txt=f'Date : {date}')

    pdf.output(f'pdf_files/{Path(filepath).stem}.pdf')

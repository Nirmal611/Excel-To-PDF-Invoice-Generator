from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

class FPDF(FPDF):
    def header(self):
        pdf.rect(x=5, y=5, w=200, h=288)

filepaths = glob.glob("Excel_Sheets/*.xlsx")

for filepath in filepaths :
    df = pd.read_excel(filepath,sheet_name='Sheet 1')
    invoice = Path(filepath).stem.split('-')[0]
    date = Path(filepath).stem.split('-')[1]

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    pdf.set_font(family='Times', size=22, style='B')
    pdf.cell(w=0, h=18, txt='Nirmal Agencies',ln=1,align='C')

    pdf.set_font(family='Helvetica', size=14, style='')
    pdf.cell(w=0, h=14, txt=f'Invoice Number : {invoice}', ln =1)
    pdf.cell(w=0, h=14, txt=f'Date : {date}', ln=1)

    col = list(df.columns)
    col = [items.replace('_',' ').title() for items in col]

    pdf.set_font(family='Courier', size=10, style='B')
    pdf.cell(w=25, h=12, txt=col[0], border=1, align='C')
    pdf.cell(w=70, h=12, txt=col[1], border=1)
    pdf.cell(w=35, h=12, txt=col[2], border=1, align='C')
    pdf.cell(w=32, h=12, txt=col[3], border=1, align='C')
    pdf.cell(w=30, h=12, txt=col[4], border=1, ln=1, align='C')

    total = 0
    for index , row in df.iterrows() :
        pdf.set_font(family='Courier', size=10, style='')
        pdf.cell(w=25,h=12,txt=str(row['product_id']),border=1, align='C')
        pdf.cell(w=70, h=12, txt=str(row['product_name']), border=1)
        pdf.cell(w=35, h=12, txt=str(row['amount_purchased']), border=1, align='C')
        pdf.cell(w=32, h=12, txt=str(row['price_per_unit']), border=1, align='C')
        pdf.cell(w=30, h=12, txt=str(row['total_price']), border=1, ln=1, align='C')
        total += int(row['total_price'])

    pdf.set_font(family='Courier', size=10, style='')
    pdf.cell(w=25, h=12, txt='', border=1, align='C')
    pdf.cell(w=70, h=12, txt='', border=1)
    pdf.cell(w=35, h=12, txt='', border=1, align='C')
    pdf.cell(w=32, h=12, txt='', border=1, align='C')
    pdf.cell(w=30, h=12, txt=str(total), border=1, ln=1, align='C')

    pdf.set_font(family='Times', size=12, style='I')
    pdf.cell(w=0, h=15, txt=f'The total due amount to be paid : {total}', border=0, ln=1, align='L')
    pdf.cell(w=0, h=15, txt='Invoice generated by Nirmal Agencies', border=0, ln=1, align='L')

    pdf.output(f'pdf_files/{Path(filepath).stem}.pdf')

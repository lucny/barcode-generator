# import EAN13 from barcode module
import os
from barcode import Code128
from barcode.writer import ImageWriter
import openpyxl
from pathlib import Path
from fpdf import FPDF

def read_excel_files():
    xlsx_files = [path for path in Path('data').rglob('*.xlsx')]
    wbs = [{'filename': wb.name,'data': openpyxl.load_workbook(wb)} for wb in xlsx_files]
    return wbs

def parse_excel_sheet(sheet):
    values = []
    for row in sheet.iter_rows():
        for cell in row:
            values.append(cell.value)
    return values

def write_code128(number):
    code = Code128(str(number), writer=ImageWriter())
    code.save(os.path.join('images/', str(number)))

def generate_barcodes(numbers):
    for n in numbers:
        write_code128(n)

def remove_files(dir, ext):
    for file in Path(dir).glob(f'*.{ext}'):
        os.remove(file)

def png_to_pdf(dir, name):
    files = Path(dir).glob('*.png')
    pdf = FPDF()
    pdf.add_page()
    r = 0
    c = 1
    for img in files:
        pdf.image(img, 10 + (c - 1) * 60, 10 + r * 40, 55, 35)
        if c % 3 != 0:
            c += 1
        else:
            c = 1
            r += 1
        if r == 7:
            pdf.add_page()
            r = 0
            c = 1
    pdf.output(os.path.join('pdfs/', f'{name}.pdf'))


if __name__ == '__main__':
    wbs = read_excel_files()
    for wb in wbs:
        numbers = parse_excel_sheet(wb['data'].active)
        print(numbers)
        generate_barcodes(numbers)
        png_to_pdf('images', wb['filename'][0:-5])
        print(f"Uložen soubor: {wb['filename'][0:-5]}.pdf")
        print("****************************************************************")
        remove_files('images', 'png')
    print("The End")
# See PyCharm help at https://www.jetbrains.com/help/pycharm/

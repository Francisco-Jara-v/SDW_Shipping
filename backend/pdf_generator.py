from docx2pdf import convert
import win32com.client
import os

def word_a_pdf(input_path, output_path):
    convert(input_path, output_path)
    
def excel_a_pdf(ruta_excel, ruta_pdf):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(os.path.abspath(ruta_excel))
    ws = wb.Worksheets(1)

    ws.ExportAsFixedFormat(0, os.path.abspath(ruta_pdf))

    wb.Close(False)
    excel.Quit()
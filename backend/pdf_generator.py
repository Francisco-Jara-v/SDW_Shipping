import win32com.client
import os


def excel_a_pdf(ruta_excel, ruta_pdf, hoja=1):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(os.path.abspath(ruta_excel))

    ws = wb.Worksheets(hoja)

    ws.ExportAsFixedFormat(
        0,
        os.path.abspath(ruta_pdf)
    )

    wb.Close(False)
    excel.Quit()
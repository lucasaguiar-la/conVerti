import os
from win32com import client

cod_dir = os.path.dirname(os.path.abspath(__file__))
arquivo_dir = os.path.join(cod_dir, "arquivo")
pdf_dir = os.path.join(cod_dir, "final.pdf")
excel = client.Dispatch("Excel.Application")

sheets = excel.Workbooks.Open(arquivo_dir)
work_sheets = sheets.Worksheets[0]

work_sheets.ExportAsFixedFormat(0, pdf_dir)

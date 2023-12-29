import os
from win32com import client


def converter_excel():
    # Caminhos
    cod_dir = os.path.dirname(os.path.abspath(__file__))
    arquivo_dir = os.path.join(cod_dir, "arquivo")
    #pdf_dir = os.path.join(cod_dir, "final.pdf")
    excel = client.Dispatch("Excel.Application")

    sheets = excel.Workbooks.Open(arquivo_dir)
    index = 0

    while True:
        try:
            pdf_dir = os.path.join(cod_dir, f"final{index}.pdf")
            work_sheets = sheets.Worksheets[index]
            work_sheets.ExportAsFixedFormat(0, pdf_dir)
            print(f"Convertendo página {index+1}")
            index += 1
        except Exception as e:
            print("Conversão .XLSX para .PDF concluída!/n")
            break

    excel.Quit()


converter_excel()
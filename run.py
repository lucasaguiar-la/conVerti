import os
from win32com import client

def converter_excel():
    # Caminhos
    cod_dir = os.path.dirname(os.path.abspath(__file__))
    arquivo_dir = os.path.join(cod_dir, "arquivo")
    #pdf_dir = os.path.join(cod_dir, "final.pdf")
    excel = client.Dispatch("Excel.Application")

    index = 0
    loop = 0
    sheets = excel.Workbooks.Open(arquivo_dir)
    output_workbook = excel.Workbooks.Add()

    while True:
        try:
            #pdf_dir = os.path.join(cod_dir, f"Arquivo_0{index}.pdf")
            work_sheets = sheets.Worksheets[index]
            work_sheets.Copy(None, output_workbook.Sheets(output_workbook.Sheets.Count))
            #work_sheets.ExportAsFixedFormat(0, pdf_dir)
            print(f"Convertendo página {index+1}")
            index += 1
        except Exception as e:
            print("Conversão .XLSX para .PDF concluída!/n")
            loop + 1
            break
    pdf_dir = os.path.join(cod_dir, f"Arquivo_0{loop}.pdf")
    output_workbook.ExportAsFixedFormat(0, pdf_dir, 0, 1) 

    output_workbook.Saved = True
    sheets.Close(False)    
    excel.Quit()

converter_excel()


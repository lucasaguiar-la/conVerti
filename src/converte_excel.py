from src.utils import definir_caminho
from win32com import client
import os

def converter_excel(aplicar_intervalo=False):
    diretorio_arquivo = definir_caminho("../arquivos/originais/excel/arquivo.xlsx")
    diretorio_destino = definir_caminho("../arquivos/convertidos/excel/")

    print("Instanciando aplicação Excel...")
    excel = client.Dispatch("Excel.Application")
    planilhas = excel.Workbooks.Open(diretorio_arquivo)

    print("Iniciando conversão...")
    if not aplicar_intervalo:
        arquivo_saida = excel.Workbooks.Add()
        for index in range(planilhas.Worksheets.Count):
            planilhas.Worksheets[index].Copy(None, arquivo_saida.Sheets(arquivo_saida.Sheets.Count))

        diretorio_pdf = os.path.join(diretorio_destino, f"ArquivoExcel.pdf")
        arquivo_saida.ExportAsFixedFormat(0, diretorio_pdf, 0, 1) 
        arquivo_saida.Close(False)

    else:
        for index, planilha in enumerate(planilhas.Worksheets, start=1):
            diretorio_pdf = os.path.join(diretorio_destino, f"ArquivoExcel_intervalo_{index}")
            planilha.ExportAsFixedFormat(0, diretorio_pdf, 0, 1)

    planilhas.Close(False)
    excel.Quit()

    return True
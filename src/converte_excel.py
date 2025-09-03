from src.utils import definir_caminho
from win32com import client
import os

def converter_excel(aplicar_intervalo=False):
    index = 0
    loop = 1
    diretorio_arquivo = definir_caminho("../arquivos/originais/excel/arquivo.xlsx")
    diretorio_destino = definir_caminho("../arquivos/convertidos/excel/")

    print("Instanciando aplicação Excel...")
    excel = client.Dispatch("Excel.Application")
    planilhas = excel.Workbooks.Open(diretorio_arquivo)
    arquivo_saida = excel.Workbooks.Add()

    print("Iniciando conversão...")
    if not aplicar_intervalo:
        while True:
            try:
                planilha_ativa = planilhas.Worksheets[index]
                planilha_ativa.Copy(None, arquivo_saida.Sheets(arquivo_saida.Sheets.Count))
                print(f"\nConvertendo planilha/sheet {index+1}")
                index += 1
            except Exception as e:
                print(f"\nAviso: {e}\n")
                print("Conversão de Excel para PDF concluída!")
                loop + 1
                break
        pdf_dir = os.path.join(diretorio_destino, f"ArquivoExcel_{loop}.pdf")
        arquivo_saida.ExportAsFixedFormat(0, pdf_dir, 0, 1) 

        arquivo_saida.Saved = True
        planilhas.Close(False)
        excel.Quit()
    else:
        print("Converter um PDF por sheet!")

    return True
from src.converte_excel import converter_excel
from ui.gui import selecionar_arquivo

if __name__ == "__main__":
    print("Selecione o arquivo Excel para conversão")
    arquivo = selecionar_arquivo()

    if not arquivo:
        print("Nenhum arquivo selecionado. Encerrando...")
    else:
        processo = converter_excel(diretorio_arquivo=arquivo, aplicar_intervalo=False)

        if processo:
            print("Processo de conversão concluído!")
        else:
            print("Algo deu errado!")

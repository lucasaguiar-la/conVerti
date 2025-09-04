from src.converte_excel import converter_excel

if __name__ == "__main__":
    print("Iniciando conversão de arquivo excel...")
    processo = converter_excel(aplicar_intervalo=False)

    if processo:
        print("Processo de conversão concluído!")
    else:
        print("Algo deu errado!")

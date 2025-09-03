from src.converte_excel import converter_excel

if __name__ == "__main__":
    print("Iniciando conversão de arquivo excel...")
    processo = converter_excel()

    if processo:
        print("Processo de conversão concluído!")

import os

def definir_caminho(nome_arquivo):
    diretorio_raiz = os.path.dirname(os.path.abspath(__file__))
    diretorio_arquivo = os.path.join(diretorio_raiz, nome_arquivo)

    return diretorio_arquivo

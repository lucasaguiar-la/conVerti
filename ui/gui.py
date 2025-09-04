from tkinter import Tk, filedialog

def selecionar_arquivo():
    root = Tk()
    root.withdraw()

    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )

    return arquivo

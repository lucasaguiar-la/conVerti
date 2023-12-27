import os
from docx import Document
from PIL import Image
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def excel_pdf(arquivo_excel, output_pdf):

    arquivo_excel = pd.read_excel("nome do arquivo")

    # Define as colunas e depois cria um gráfico, para ser salvo posteriormente (get_figure)
    x_column, y_column = arquivo_excel.columns[:10]
    plot = arquivo_excel.plot(x=x_column, y=y_column, kind='bar').get_figure()

    # Buffer para lermos o gráfico criado e transforma-lo em um png
    buffer = BytesIO()
    plot.savefig(buffer, format='png')
    buffer.seek(0)

    # Processo para salvar o PDF através da imagem 'desenhada' pelo canvas
    c = canvas.Canvas(output_pdf, pagesize=letter)
    c.drawImage(buffer, 100, 400, width=400, height=300)
    c.showPage()
    c.save()

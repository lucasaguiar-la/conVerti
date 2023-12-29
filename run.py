import os
import pandas as pd
from PIL import Image
from io import BytesIO
from docx import Document
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader

def excel_pdf(arquivo_excel, output_pdf):

    arquivo_excel = pd.read_excel("nome do arquivo.xlsx")

    c = canvas.Canvas(output_pdf, pagesize=letter)

    x, y = 100, 800

    for _, row in arquivo_excel.iterrows():
        for col_name, value in row.items():
            c.drawString(x, y, f"{col_name}:")
            x += 80

            c.drawString(x, y, str(value))
            x += 120

        y -= 20
        x = 100

    c.showPage()
    c.save()

excel_pdf("nome do arquivo.xlsx", "teste.pdf")
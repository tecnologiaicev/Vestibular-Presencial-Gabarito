from pdfminer.high_level import extract_text
from pdfminer.layout import LAParams

def extrair_info_layout(pdf_path):
    with open(pdf_path, 'rb') as file:
        text = extract_text(file, laparams=LAParams())

    return text

# Substitua 'Carderno-vest.pdf' pelo seu PDF existente
pdf_path = 'teste_caderno.pdf'

# Extrair informações de layout do PDF
layout_info = extrair_info_layout(pdf_path)

# Imprimir as informações do layout
print(layout_info)

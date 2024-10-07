import PyPDF2
from reportlab.pdfgen import canvas
import io
from reportlab.lib.units import inch
import pandas as pd
import os

def preencher_pdf(pdf_entrada, pdf_saida, nome, curso, id_aluno, cpf, i1, i2, i3, i4, i5, tamanho_fonte_nome=12, tamanho_fonte_curso=12, tamanho_fonte_id=18, tamanho_fonte_cpf=10):
    with open(pdf_entrada, 'rb') as file:
        reader = PyPDF2.PdfReader(file)

        # Criar um objeto PdfWriter
        writer = PyPDF2.PdfWriter()

        # Obter a primeira página do PDF 
        page = reader.pages[0]

        # Converter o nome do aluno para maiúsculas
        nome_aluno = nome.upper()

        # Adicionar nome
        nova_pagina_nome = create_text_page(nome_aluno, 125, 745, tamanho_fonte_nome)
        page.merge_page(nova_pagina_nome)

        # Adicionar curso
        nova_pagina_curso = create_text_page(curso, 50, 725, tamanho_fonte_curso)
        page.merge_page(nova_pagina_curso)

        # Adicionar ID do aluno
        nova_pagina_id = create_text_page(str(id_aluno), 475, 715, tamanho_fonte_id)
        page.merge_page(nova_pagina_id)

        # Ajustar o formato do CPF
        cpf_formatado = f"{int(cpf):011d}"  # Garante que o CPF tenha 11 dígitos e preenche com zeros à esquerda
        cpf_formatado = f"{cpf_formatado[:3]}.{cpf_formatado[3:6]}.{cpf_formatado[6:9]}-{cpf_formatado[9:]}"

        # Adicionar CPF
        nova_pagina_cpf = create_text_page(cpf_formatado, 145, 699, tamanho_fonte_cpf)
        page.merge_page(nova_pagina_cpf)
        
        # Adicionar ids
        nova_pagina_cpf = create_text_page(str(i1), 345, 320, tamanho_fonte_cpf)
        page.merge_page(nova_pagina_cpf)
        # Adicionar ids
        nova_pagina_cpf = create_text_page(str(i2), 365, 320, tamanho_fonte_cpf)
        page.merge_page(nova_pagina_cpf)
        # Adicionar ids
        nova_pagina_cpf = create_text_page(str(i3), 385, 320, tamanho_fonte_cpf)
        page.merge_page(nova_pagina_cpf)
        # Adicionar ids
        nova_pagina_cpf = create_text_page(str(i4), 405, 320, tamanho_fonte_cpf)
        page.merge_page(nova_pagina_cpf)
        # Adicionar ids
        nova_pagina_cpf = create_text_page(str(i5), 425, 320, tamanho_fonte_cpf)
        page.merge_page(nova_pagina_cpf)
          ######################################################  ######################################################
        # Adicionar página modificada ao escritor
        # Criar uma nova página com os retângulos
        nova_pagina_retangulos = create_circles_page([i1, i2, i3, i4, i5])

        # Mesclar a página dos retângulos com a página original
        page.merge_page(nova_pagina_retangulos)

        # Adicionar página modificada ao escritor
        writer.add_page(page)

        # Salvar o PDF modificado
        with open(pdf_saida, 'wb') as output_file:
            writer.write(output_file)

def create_text_page(text, x, y, font_size):
    packet = io.BytesIO()
    can = canvas.Canvas(packet)
    can.setFont("Helvetica", font_size)
    can.drawString(x, y, str(text))
    can.save()

    packet.seek(0)
    new_pdf = PyPDF2.PdfReader(packet)
    return new_pdf.pages[0]

def create_circles_page(ids):
    packet = io.BytesIO()
    can = canvas.Canvas(packet)

    fill_color = (0, 0, 0)

    colunas = 5
    linhas = 10
    raio_circulo = 8  # Tamanho do raio do círculo
    espacamento_x = 5  # Ajuste fino na posição horizontal
    espacamento_y = 6  # Ajuste fino na posição vertical

    for i, id_valor in enumerate(ids):
        if id_valor > 5:
            linha = (linhas - 1) - (id_valor % linhas)  # Inverter a linha
            ajuste_horizontal = 0  # Ajuste fino na posição horizontal para círculos não nulos
            coluna = i  # Usar divisão inteira para obter a coluna correta

            x = 336 + coluna * (2 * raio_circulo + espacamento_x) + ajuste_horizontal
            y = 85 + linha * (2 * raio_circulo + espacamento_y)
        else:
            
            linha = (linhas) - (id_valor % linhas)  # Inverter a linha
            ajuste_horizontal = 2  # Ajuste fino na posição horizontal para círculos nulos

            coluna = i  # Usar divisão inteira para obter a coluna correta

            x = 335 + coluna * (2 * raio_circulo + espacamento_x) + ajuste_horizontal
            y = 69 + linha * (2 * raio_circulo + espacamento_y)

        can.setFillColorRGB(*fill_color)
        can.circle(x + raio_circulo, y + raio_circulo, raio_circulo, fill=1)

    can.save()

    packet.seek(0)
    new_pdf = PyPDF2.PdfReader(packet)
    return new_pdf.pages[0]








# Substitua 'Carderno-vest.pdf' pelo seu PDF existente
pdf_entrada = 'Carderno-vest.pdf'
df = pd.read_excel('vest.xlsx')  # Substitua 'dados_alunos.xlsx' pelo nome do seu arquivo Excel
pasta_destino = 'gabarito/'

# Certifique-se de que o diretório de destino exista, se não existir, crie-o
if not os.path.exists(pasta_destino):
    os.makedirs(pasta_destino)
# Iterar sobre as linhas do DataFrame
for index, row in df.iterrows():
    # Extrair informações de cada linha
    nome_aluno = row['Nome']
    curso_aluno = row['Curso']
    id_aluno = row['ID']
    i1 = row['i1']
    i2 = row['i2']
    i3 = row['i3']
    i4 = row['i4']
    i5 = row['i5']
    cpf_aluno = row['CPF']

    # Modificar PDF e salvar
    print(nome_aluno)
    pdf_saida = os.path.join(pasta_destino, f"{nome_aluno}_caderno.pdf")
    contador = 1    
    while os.path.exists(pdf_saida):
        pdf_saida = os.path.join(pasta_destino, f"{nome_aluno}_caderno_{contador}.pdf")
        contador += 1
    preencher_pdf(pdf_entrada, pdf_saida, nome_aluno, curso_aluno, id_aluno, cpf_aluno, i1, i2, i3, i4, i5)

print("Processo concluído.")
import openpyxl as xl
from PIL import Image, ImageDraw, ImageFont
import tkinter as tk
from tkinter import filedialog

fonte_geral = filedialog.askopenfilename(title="Fonte Geral", filetypes=[("Font files", ".ttf")]) 
fonte_geral_caminho = filedialog.askopenfilename(title="Fonte Geral", filetypes=[("Font files", ".ttf")]) 

fonte_nome = filedialog.askopenfilename(title="Fonte para os nomes", filetypes=[("Font files", ".ttf")])
fonte_nome_caminho = filedialog.askopenfilename(title="Fonte para os nomes", filetypes=[("Font files", ".ttf")])
 

planilha = filedialog.askopenfilename(title="Planilha Base" ,filetypes=[("Excel files", ".xlsx .xls")])

imagem_base = filedialog.askopenfilename(title="Imagem Base" ,filetypes=[("Image files", ".jpg .png .jpeg")])

workbook_alunos = xl.load_workbook(planilha)
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_aluno = linha[1].value
    tipo_aluno = linha[2].value
    carga_horaria = linha[5].value
    
    data_inicio = linha[3].value
    data_final = linha[4].value
    data_emissao = linha[6].value
    
    fonte__nome = ImageFont.truetype(fonte_nome_caminho,90)
    fonte_geral = ImageFont.truetype(fonte_geral_caminho,80)
    fonte_data = ImageFont.truetype(fonte_geral_caminho,55)
    
    image = Image.open(imagem_base)
    desenhar = ImageDraw.Draw(image)
    
    desenhar.text((1020,827), nome_aluno, fill='black', font=fonte__nome)
    desenhar.text((1060,950), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1435,1065), tipo_aluno, fill='black', font=fonte_geral)
    desenhar.text((1480,1182), str(carga_horaria), fill='black', font=fonte_geral)
    
    desenhar.text((750, 1770), data_inicio, fill='blue', font=fonte_data)
    desenhar.text((750, 1930), data_inicio, fill='blue', font=fonte_data)
    
    desenhar.text((2220, 1930), data_emissao, fill='blue', font=fonte_data)
    
    image.save(f'./Certificados Prontos/{indice} {nome_aluno}certificado.png')
# coding: utf8

import xlrd
import pandas as pd
import numpy as nd
import openpyxl
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader


import os

arquivos_prontuarios = './Prontuários'


#leitura dos arquivos excel para importação, mapeamento de células
for diretorio, subpastas, arquivos in os.walk(arquivos_prontuarios):
    for arquivo in arquivos: 
        dados_prontuario = xlrd.open_workbook(os.path.join(diretorio,arquivo))
        aba_atendimento = dados_prontuario.sheet_by_index(0)       
        for row_num in range (aba_atendimento.nrows):
            if row_num ==0:
                continue
            row = aba_atendimento.row_values(row_num)
            print(row) #dados arquivo
        print(os.path.join(diretorio,arquivo)) #nome do arquivo o qual os dados foram extraídos

pxl_doc = openpyxl.load_workbook('./Prontuários/Prontuário_BC106_BR002.xlsx')
sheet = pxl_doc['Modelo em branco']
#calling the image_loader
image_loader = SheetImageLoader(sheet)

#get the image (put the cell you need instead of 'A1')
image = image_loader.get('E4')

#showing the image
image.show()
#ALTERAR ESTRUTURA DO CÓDIGO PARA ALINHAMENTO DE LEITURA DE ARQUIVO 
#CONFORME ENTRADA DE DADOS. 
#salvar foto animal tipo_blob bd fauna.
#saving the image
image.save('./Prontuários/Teste/image_name.jpg')
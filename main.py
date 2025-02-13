import openpyxl
from openpyxl.styles import Alignment, Font
import subprocess
import pandas as pd

file = "C:/git/linea-ordenacao-requisicao/Lista_Map.CSV"


df= pd.read_csv(file, delimiter=';', encoding='latin1')


wb = openpyxl.load_workbook('layout.xlsx')
ws = wb.active
ws.title = "Tabela"

codigo = input("Digite o Código do Produto: ")
quantidade = 

#--Produto
ws["B2"] = """Parafuso 3,5X12 CAB FLANGEADA"""
cell = ws["B2"]
cell.font = Font(size=60)
cell.alignment = Alignment(wrapText=True,horizontal='center', vertical='center')

#--Quantidade
ws["B3"] = "11130585"
cell = ws["B3"]
cell.font = Font(size=100)
cell.alignment = Alignment(wrapText=True,horizontal='center', vertical='center')


#--Numero Pedido
ws["B4"] = "6016"
cell = ws["B4"]
cell.font = Font(size=150)
cell.alignment = Alignment(wrapText=True,horizontal='center', vertical='center')


#--Volume
ws["B5"] = "10"
cell = ws["B5"]
cell.font = Font(size=110)
cell.alignment = Alignment(wrapText=True,horizontal='center', vertical='center')

for cell in ws[1]:
    cell.alignment = Alignment(horizontal='center')

# Salve o arquivo Excel
wb.save('tabela_imprimir.xlsx')

# Abra o arquivo Excel salvo
subprocess.Popen(['start', 'excel.exe', 'tabela_imprimir.xlsx'], shell=True)

print("Arquivo Excel aberto com sucesso!")
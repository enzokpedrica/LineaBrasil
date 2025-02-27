import pandas as pd
import openpyxl
import subprocess

colunas_esperadas = ['Produto_x']

file = "Planilha.xlsx"
df = pd.read_csv(file, delimiter=';', encoding='latin1', names=colunas_esperadas, skiprows=1)

# print(df[['Produto','Descrição Produto','Qtde Mov']])

file_map = "Lista_Map.CSV"
df_map = pd.read_csv(file_map, delimiter=';', encoding='latin1')

df_merged = pd.merge(df, df_map, left_on='Produto', right_on='Cod Produto', how="inner")

# Ordenar e selecionar as colunas necessárias
df_merged = df_merged.sort_values(by='Secao')
df_merged = df_merged[['Produto_x', 'Descrição Produto', 'Qtde Mov', 'Rua', 'Secao', 'Andar']]
df_merged = df_merged.drop_duplicates(subset=['Descrição Produto'])

df_merged.to_excel('tabela_imprimir.xlsx', index=False, engine='openpyxl')

wb = openpyxl.load_workbook('tabela_imprimir.xlsx')
ws = wb.active
ws.title = "Tabela"

# Salve o arquivo Excel
wb.save('tabela_imprimir.xlsx')

# Abra o arquivo Excel salvo
subprocess.Popen(['start', 'excel.exe', 'tabela_imprimir.xlsx'], shell=True)


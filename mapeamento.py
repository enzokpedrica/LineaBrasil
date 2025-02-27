import pandas as pd
import openpyxl
import subprocess

file = "Planilha.xlsx"
df = pd.read_excel(file, engine='openpyxl')

file_map = "Lista_Map.CSV"
df_map = pd.read_csv(file_map, delimiter=';', encoding='latin1')

# Fazer o merge (join) entre os DataFrames, mantendo apenas os produtos que estão no Lista_Map
df_merged = pd.merge(df, df_map, left_on='Produto', right_on='Cod Produto', how="inner")

# Verificar se apenas os produtos mapeados estão no DataFrame resultante
df_merged = df_merged.dropna(subset=['Cod Produto'])

# Ordenar e selecionar as colunas necessárias
df_merged = df_merged.sort_values(by='Secao')
df_merged = df_merged[['Produto_x', 'Descrição Produto', 'Qtde Mov', 'Rua', 'Secao', 'Andar']]
df_merged = df_merged.drop_duplicates(subset=['Descrição Produto'])

# Salvar o resultado em um novo arquivo Excel
df_merged.to_excel('tabela_imprimir.xlsx', index=False, engine='openpyxl')

# Renomear a aba ativa
wb = openpyxl.load_workbook('tabela_imprimir.xlsx')
ws = wb.active
ws.title = "Tabela"
wb.save('tabela_imprimir.xlsx')

# Abrir o arquivo Excel salvo
subprocess.Popen(['start', 'excel.exe', 'tabela_imprimir.xlsx'], shell=True)

print("Join concluído e arquivo tabela_imprimir.xlsx salvo.")



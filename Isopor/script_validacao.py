import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.load_workbook("teste.xlsx")
sheet = wb.active

validacao = '"Mapeado, Conferido, Entregue"'

rule = DataValidation(type='list',formula1=validacao, allow_blank=True)

rule.error  = "Entrada Inválida"
rule.errorTitle = "Selecione Corretamente"

rule.prompt = "Selecione a Lista"
rule.promptTitle = "Selecione a opção"

sheet.add_data_validation(rule)

rule.add("C2:C100")

# Salva o novo arquivo
wb.save("EKP_Isopores.xlsx")
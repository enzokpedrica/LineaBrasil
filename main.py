from Mapeamento import mapeamento as mp
from Ordenar_Requisicao	import Ordenador as od


print("--------------------------------")
print("Seja Bem Vindo ao EKP_TOOLS")
print("--------------------------------")
print("Selecione a opção que necessita:")
resposta = int(input(" 1 - Mapeamento para Armazenagem \n 2 - Ordenador de Requisição do Acessórios \n 3 - Sair \n Digite a opção: "))

if resposta == 1:
    print(" ")
    print("O Arquivo foi salvo corretamente?")
    confirmacao = int(input(" 1 - Sim \n 2 - Não \n Digite a opção: "))
    if confirmacao == 1:
        mp.mapeamentoItens()
    else:
        print("Por favor salve o Arquivo corretamente")
        
elif resposta == 2:
    print(" ")
    print("O Arquivo foi salvo corretamente?")
    confirmacao = int(input(" 1 - Sim \n 2 - Não \n Digite a opção: "))
    if confirmacao == 1:
            od.requisicaoAcessorios()
    else:
        print("Por favor salve o Arquivo corretamente")
else:
    print("Até breve")





### Opções: DEF(defeito), MLH(melhoria), NOV(nova funcionalidade), DIS(discussão)
### Status: INI(iniciado), ABO(abortado), CON(concluído)
## 11-02-2025
(DIS) Analisar como substituir CSV por algum banco de dados permitindo acesso de diferentes localidades

## 12-02-2025
(NOV) Adicionar um sistema de alteração da localização do Pallet
(DEF) Notei um erro, o código do produto pode estar presente em mais de 1 localização logo se a requisição contém um produto onde o mesmo possui 2 localizações o pedido é duplicado. (ex: 160006033 - Parafuso - R3E - S3 - E0 - 100000.00 / -- / 160006033 - Parafuso - R3E - S3 - E2 - 100000.00) (CON)

## 27-02-2025
(NOV) Organizar a criação de um mapeamento dos itens que chegaram no dia
(MLH) Deixar os arquivos mais organizador, criar módulos e deixar o main.py sendo o principal
(DEF) Um erro que está acontecendo é no DropDuplicates, porque pode ocorrer de o mesmo código ser produtos diferentes, porque o que vai diferenciar um produto do outro é a Deriv, que seria basicamente a cor de cada item
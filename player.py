from lowModel.main import *

planilha = Excel(r'C:\Users\20221214010004\Desktop\Python\a.xlsx')

tabelaDados = planilha.getTableCells('Planilha1',headerRow=4)

tabelaNova = planilha.getTableCells('Planilha1',headerRow=19)

tabelaComparacao = planilha.getTableCells('Planilha1',headerRow=31)

print('Tabela de Dados:')
for coluna in tabelaDados.columns:
    print(f'    Coluna {coluna.name}:')
    print(f'        Começa na linha: {coluna.firstRow}')
    print(f'        Termina na linha: {coluna.lastRow}')



print('\n\n\nTabela Nova:')
for coluna in tabelaNova.columns:
    print(f'    Coluna {coluna.name}:')
    print(f'        Começa na linha: {coluna.firstRow}')
    print(f'        Termina na linha: {coluna.lastRow}')

print('\n\n\nTabela de Comparação:')
for coluna in tabelaComparacao.columns:
    print(f'    Coluna {coluna.name}:')
    print(f'        Começa na linha: {coluna.firstRow}')
    print(f'        Termina na linha: {coluna.lastRow}')

planilha.close()
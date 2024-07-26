from lowModel.main import *

planilha = Excel(r'C:\Users\lgpfr\Desktop\GU\Programação\Python\a.xlsx')

table = planilha.getTableCells('Planilha1',4)

for column in table.columns:
    print(f'-------{column.name}:')
    for cell in column.values:
        print(cell.row)

planilha.close()
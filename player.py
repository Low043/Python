from lowModel.main import *

planilha = Excel(r'C:\Users\lgpfr\Desktop\GU\Programação\Python\a.xlsx')

tabelaDados = planilha.getTable('Planilha1',headerRow=4)

tabelaNova = planilha.getTable('Planilha1',headerRow=19)

tabelaNova.pullValuesFrom(tabelaDados)

planilha.save()

planilha.close()
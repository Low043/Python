from lowModel.main import *
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

planilhaPedro = Excel(r'C:\Users\luisg\Desktop\classificação\pedro.xlsx')
tabelaPedro = planilhaPedro.getTable('CLASSIFICAÇÃO SÍNTESE',headerRow=9)

for i in range(1,8):
    mes = numToMonth(i)
    planilhaMes = Excel(rf'C:\Users\luisg\Desktop\classificação\planilhas\{i}. {mes}\CLASSIFICAÇÃO (PRODUÇÃO) - {mes} 2024.xlsx')
    tabelaMes = planilhaMes.getTable('CLASSIFICAÇÃO ',headerRow=6)
    tabelaPedro.pullValuesFrom(tabelaMes)
    planilhaMes.close()

planilhaPedro.save()
planilhaPedro.close()
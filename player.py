from lowModel.main import *

excel = Excel(r'C:\Users\lgpfr\Desktop\GU\Programação\Python\a.xlsx')

r = excel.getCellRange('Plan1','A',1,'C',5)
excel.setCellRange('Plan1','D',6,'F',10,r)

excel.save()
excel.close()
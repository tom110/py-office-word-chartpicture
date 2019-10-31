import xlrd
from xlutils.copy import copy

excel_path = "C:\\Users\\Administrator\\Desktop\\需要调整表格\\需要调整表格\\1.地表.xls"
data = xlrd.open_workbook(excel_path, formatting_info=True)

exeExcel = copy(data)

exeExcel.save("outtest.xls")
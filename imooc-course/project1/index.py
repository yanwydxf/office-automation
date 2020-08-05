import xlrd

data=xlrd.open_workbook("data1.xlsx")
# print(data.sheet_loaded(0))
# data.unload_sheet(0)
# print(data.sheet_loaded(0))
# print(data.sheet_loaded(1))
# print(data.sheets())#获取全部sheet
# print(data.sheets()[0])
# print(data.sheet_by_index(0))#根据索引获取工作表
# print(data.sheet_by_name("Sheet1"))#g根据sheetname进行获取
# print(data.sheet_names())#获取所有工作表的name
# print(data.nsheets)#返回excle工作表的数量

#操作excel行
sheet=data.sheet_by_index(1)#获取第一个工作表
# print(sheet.nrows) #获取sheet下的有效行数
# print(sheet.row(1))#该行单元格对象组成的列表
print(sheet.row_types(2,4))#获取单元格的数据类型
# print(sheet.row(1)[2].value) #得到单元格value
# print(sheet.row_values(1))#得到指定行单元格的value
# print(sheet.row_len(1))#得到单元格的长度

#操作excel列
# sheet=data.sheet_by_index(0)
# print(sheet.ncols)
# print(sheet.col(1))#该列单元格对象组成的列表
# print(sheet.col(1)[2].value)
# print(sheet.col_values(1))#返回该列所有单元格的value组成的列表
# print(sheet.col_types(5))

#操作Excel单元格
# sheet=data.sheet_by_index(0)
# print(sheet.cell(1,2))
# print(sheet.cell_type(1,2))
# print(sheet.cell(1,2).ctype)#获取单元格数据类型
# print(sheet.cell(1,2).value)
# print(sheet.cell_value(1,2))
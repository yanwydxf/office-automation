import xlsxwriter

wb = xlsxwriter.Workbook("data.xlsx")
cell_format = wb.add_format({'bold': True})  # 格式对象
cell_format1 = wb.add_format()  # 格式对象
cell_format1.set_bold()  # 设置加粗
cell_format1.set_font_color("red")  # 文本的颜色
cell_format1.set_font_size(14)
cell_format1.set_align("center")  # 对齐方式
cell_format2 = wb.add_format()  # 格式对象
cell_format2.set_bg_color("#808080")  # 设置背景颜色
# 创建sheet
sheet = wb.add_worksheet("newsheet")
# 写入
# sheet.write_string()
sheet.write(0, 0, "2020年度", cell_format)  # 写入单元格数据
sheet.merge_range(1, 0, 2, 2, "第一季度销售统计", cell_format1)
data = (
    ["一月份", 500, 450],
    ["二月份", 600, 450],
    ["三月份", 700, 550]
)
sheet.write_row(3, 0, ["月份", "预期销售额", "实际的销售额"], cell_format2)
# 写入数据
for index, item in enumerate(data):
    sheet.write_row(index + 4, 0, item)
# 写入excle公式
sheet.write(7, 1, "=SUM(B5:B7)")
sheet.write(7, 2, '=SUM(C5:C7)')
sheet.write_url(9, 0, "http://www.baidu.com", string="更多数据")
sheet.insert_image(10, 0, "view.png")  # 插入图片

#写入
chart=wb.add_chart({'type':'line'})
chart.set_title({'name':'第一季度销售统计'})
#X Y 描述信息
chart.set_x_axis({'name':'月份'})
chart.set_y_axis({'name':'销售额'})
#数据
chart.add_series({
    'name':'预期销售额',
    'categories':'=newsheet!$A$5:$A$7',
    'values':['newsheet',4,1,6,1]
})
chart.add_series({
    'name':'实际销售额',
    'categories':'=newsheet!$A$5:$A$7',
    'values':['newsheet',4,2,6,2],
    'data_labels':{'value':True}
})
sheet.insert_chart('A23',chart)
wb.close()

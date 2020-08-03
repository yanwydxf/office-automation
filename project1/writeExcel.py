import xlwt

titlestyle = xlwt.XFStyle()  # 初始化样式
titlefont = xlwt.Font()
titlefont.name = "宋体"
titlefont.bold = True  # 加粗
titlefont.height = 11*20  # 字号
titlefont.colour_index = 0x08  # 字体颜色
titlestyle.font = titlefont

# 单元格对齐方式
cellalign = xlwt.Alignment()
cellalign.horz = 0x02
cellalign.vert = 0x01
titlestyle.alignment = cellalign

# 边框
borders = xlwt.Borders()
borders.right = xlwt.Borders.DASHED
borders.bottom = xlwt.Borders.DOTTED
titlestyle.borders = borders

# 背景颜色
datestyle = xlwt.XFStyle()
bgcolor = xlwt.Pattern()
bgcolor.pattern=xlwt.Pattern.SOLID_PATTERN
bgcolor.pattern_fore_colour=22 #背景颜色
datestyle.pattern=bgcolor
# 第一步：创建工作簿
wb = xlwt.Workbook()
# 第二步：创建工作表
ws = wb.add_sheet("CNY")
# 第三步：填充数据
ws.write_merge(0, 1, 0, 5, "2019年货币兑换表", titlestyle)

# 写入货币数据
data = (("Date", "英镑", "人民币", "港币", "日元", "美元"
         ), ("01/01/2019", 8.722551, 1, 0.877885, 0.062722, 6.8759
             ), ("02/01/2019", 8.634922, 1, 0.875731, 0.062773, 6.8601
                 ))
for i, item in enumerate(data):
    for j, val in enumerate(item):
        if j==0:
            ws.write(i + 2, j, val,datestyle)
        else:
            ws.write(i + 2, j, val)
# 创建第二个工作表

wsimage = wb.add_sheet("image")
# 写入图片
wsimage.insert_bitmap("2017.bmp", 0, 0)
# 第四步：保存
wb.save("2019-CNY.xls")

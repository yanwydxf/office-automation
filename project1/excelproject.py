import xlrd

data=xlrd.open_workbook("data2.xlsx")
sheet=data.sheet_by_index(0) #获取到工作表
questionList=[] #构建试题列表

#试题类
class Question:
    pass
for i in range(sheet.nrows):
    if i>1:
        obj=Question() #构建试题对象
        obj.subject=sheet.cell(i,1).value #题目
        obj.questionType = sheet.cell(i, 2).value  #题型
        obj.optionA = sheet.cell(i, 3).value  # 选项a
        obj.optionB = sheet.cell(i, 4).value  # 选项b
        obj.optionC = sheet.cell(i, 5).value  # 选项c
        obj.optionD = sheet.cell(i, 6).value  # 选项D
        obj.score = sheet.cell(i, 7).value  # 分值
        obj.answer = sheet.cell(i, 8).value  # 正确答案
        questionList.append(obj)

print(questionList)
#导入操作 pymysql pip install
from mysqlhelper import *
#1.链接到数据
db=dbhelper('127.0.0.1',3306,"root","123456","test")
#插入语句
sql="insert into question(subject,questionType,optionA,optionB,optionC,optionD,score,answer) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"
val=[]#空列表来存储元组
for item in questionList:
    val.append((item.subject,item.questionType,item.optionA,item.optionB,item.optionC,item.optionD,item.score,item.answer))
# print(val)
db.executemanydata(sql,val)

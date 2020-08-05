# import zmail
# 获取最新邮件并打印邮件信息
# server = zmail.server('389818529@qq.com', 'mdgxgiwpnkspbigi')
# mail = server.get_latest()
# zmail.show(mail)
# print(mail["id"])
# print(mail["from"])
# print(mail["to"])
# print(mail["subject"])
# print(mail["context_text"])
# print(mail["context_html"])

import zmail
import pymsgbox
server = zmail.server('389818529@qq.com', 'mdgxgiwpnkspbigi')
mail = server.get_latest()
mail_id = mail["id"]
# pymsgbox.alert(mail_id) #弹窗显示id

old_mailid = open('id.txt', 'r').readline()
# 写入新邮件ID
with open('id.txt', mode='w+', encoding='utf-8') as f:
    f.write(str(mail_id))
    
#判断邮件是否是最新
if old_mailid != str(mail_id):
    pymsgbox.alert("你有一封新邮件！")  # 弹窗提示有新邮件

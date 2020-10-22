import excel_split
import send_email
from settings import source_excel, host, port, user, password


# 第一步 拆分excel
excel_split.split(source_excel)


# 第二步 发送邮件(先编辑好通讯录)
# send_email.send(host, port, user, password)

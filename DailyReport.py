# creator = wangkai
# creation time = 2019/4/30 21:37
# version 1.1

import openpyxl
import time
from config import *
import smtplib
from email.mime.multipart import MIMEMultipart, MIMEBase
from email import encoders
import os
import sys

work = None
sheet = None

def init():
    """初始化"""
    global work, sheet
    date = time.strftime("%Y%m%d", time.localtime())
    date2 = time.strftime("%Y.%m.%d", time.localtime())
    print("today:", date)
    rename = "个人工作日志{}{}.xlsx".format(USER_NAME, date)
    work = openpyxl.load_workbook("个人工作日志.xlsx")
    sheet = work["Sheet1"]
    sheet["A2"] = "                                         制定日期：{}			".format(date2)
    return rename

def input_info(name):
    """输入日报，保存之"""
    morning_data = input("上午：")
    sheet["D3"] = morning_data
    afternoon_data = input("下午：")
    sheet["D4"] = afternoon_data
    work.save(name)
    work.close()

def send(name):
    """发送邮件"""
    # 配置邮件信息
    message = MIMEMultipart()
    message['from'] = SENDER
    message['to'] = RECEIVE[0]
    message['Cc'] = ",".join(ACC)
    date = time.strftime("%Y%m%d", time.localtime())
    subject = "日报-{}-{}".format(USER_NAME, date)
    message['Subject'] = subject

    # 构造附件 1
    # att = MIMEText(open(name, 'rb').read(), 'base64', 'utf-8')
    # att["Content-Type"] = 'application/octet-stream'
    # # att['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    # # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
    # att["Content-Disposition"] = 'attachment; filename="'+str(name.encode("gb2312"))+'"'
    # message.attach(att)

    # 构造附件 2 解决附件名有中文问题
    att = MIMEBase('application', 'octet-stream')
    att.set_payload(open(name, 'rb').read())
    att.add_header('Content-Disposition', 'attachment', filename=('gbk', '', name))
    encoders.encode_base64(att)
    message.attach(att)

    # 邮件发送 目前只支持qq和163
    if "163" in  MAIL_HOST:
        smtp_obj = smtplib.SMTP()
    elif "qq" in MAIL_HOST:
        smtp_obj = smtplib.SMTP_SSL(MAIL_HOST, MAIL_PORT)  # 启用SSL发信
    else:
        print("Error: 没有改邮箱主机")
        os.remove(name)
        sys.exit()

    try:
        if "163" in MAIL_HOST:
            smtp_obj.connect(MAIL_HOST, port=MAIL_PORT)
        smtp_obj.login(SENDER, MALL_PASS)  # 登录验证
        smtp_obj.sendmail(SENDER, RECEIVE+ACC , message.as_string())  # 发送
        print("邮件发送成功 ^_^")
        smtp_obj.quit()
    except smtplib.SMTPException as e:
        print("Error:邮件发送失败 @_@")
        print(e)
    finally:
        smtp_obj.close()


if __name__ == "__main__":
    print("<<< 轻松日报 1.1 Made by Wo-ki >>>")

    if USER_NAME == "" or SENDER == "" or RECEIVE[0] == "" or MALL_PASS == "" or ACC[0] == "" or ACC[1] == "":
        print("Error:config.py 未完成配置，请配置！！！")
        sys.exit()
    new_file_name = init()
    input_info(new_file_name)
    r = input("发送邮件？(Y/n):")
    while r != 'Y' and r != 'y' and r != 'N' and r != 'n':
        r = input("请输入Y或者n（不区分大小写）:")
    if r == "Y" or r == "y":
        send(new_file_name)
    else:
        print("取消发送邮件")
    # os.remove(new_file_name)

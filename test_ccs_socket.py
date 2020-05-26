import datetime
import socket
import smtplib
from email.mime.text import MIMEText
from email.header import Header
##话机登陆
def phone_login_test(conn):
    data = '003610141000010226        '
    "003610141000010226"
    conn.send(data.encode())
    recieved = conn.recv(1024)
    if recieved.decode()[0] == 1:
        return True
    else:
        return False
###学生卡登陆
def student_login_test(conn):
    data="0929548133               00010226          0303236537                          20190328143711"
    conn.send(data.encode())
    recieved = conn.recv(1024)
    if recieved.decode()[0] == 1:
        return True
    else:
        return False

def send_email(text):
    email_from = "284693929@qq.com"  # 改为自己的发送邮箱
    email_to = "liuqimin2009@163.com"  # 接收邮箱
    hostname = "smtp.qq.com"  # 不变，QQ邮箱的smtp服务器地址
    login = "284693929@qq.com"  # 发送邮箱的用户名
    password = "xddflpwqesfkbidf"  # 发送邮箱的密码，即开启smtp服务得到的授权码。注：不是QQ密码。
    subject = "亲情电话错误日志"  # 邮件主题
    smtp = smtplib.SMTP_SSL(hostname)  # SMTP_SSL默认使用465端口
    smtp.login(login, password)

    msg = MIMEText(text, "plain", "utf-8")
    msg["Subject"] = Header(subject, "utf-8")
    msg["from"] = email_from
    msg["to"] = email_to

    smtp.sendmail(email_from, email_to, msg.as_string())
    smtp.quit()


import socket
ip='192.168.1.104'
ip2='47.110.139.213'
port=7070
sock = socket.socket()
sock.connect((ip2, port))
data="0010051412"   ###心跳包
sock.sendall(data.encode())
recieved = sock.recv(1024)
print("接收到",recieved.decode())
if recieved.decode()!=data:
    print("接收到错误心跳")
sock.close()


print('开始循环!')
i=0
while(1):
    conn,addr=sock.accept()
    ###心跳包测试
    if not send_heart_package(conn):
        send_email('心跳包测试失败！')
    else:
        print("心跳包测试通过！")
    # # 公话认证：
    # if not phone_login_test(conn, 3):
    #     send_email('公话认证失败！')
    #     conn.close()
    #     break
    #
    # if not student_login_test(conn, 3):
    #     send_email('学生登陆失败！')
    #     conn.close()
    #     break
    conn.close()
    break

import smtplib
import email
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
import time
import os
from datetime import datetime

cur_path = os.path.dirname(os.path.realpath(__file__))  # 当前项目路径
# cur_path = os.path.dirname(sys.executable)  # 打包运行路径
log_path = cur_path + '\\logs'  # log_path为存放日志的路径
if not os.path.exists(log_path): os.mkdir(log_path)  # 若不存在logs文件夹，则自动创建


def File_Read(file_path):
    global Lines
    Lines = []
    with open(file_path, 'r') as f:
        while True:
            line = f.readline()  # 逐行读取
            if not line:  # 到 EOF，返回空字符串，则终止循环
                break
            File_Data(line, flag=1)
    File_Data(line, flag=0)


def File_Data(line, flag):
    global Lines
    if flag == 1:
        Lines.append(line)
    else:
        return Lines


class Log:

    def __init__(self):
        now_time = datetime.now().strftime('%Y-%m-%d')
        self.__all_log_path = os.path.join(log_path, now_time + "-all" + ".log")  # 收集所有日志信息文件
        self.__error_log_path = os.path.join(log_path, now_time + "-error" + ".log")  # 收集错误日志信息文件
        self.__send_error_log_path = os.path.join(log_path, now_time + "-send_error" + ".log")  # 收集发送失败邮箱信息文件
        self.__send_done_log_path = os.path.join(log_path, now_time + "-send_done" + ".log")  # 收集发送成功邮箱信息文件

    def SaveAllLog(self, message):
        with open(r"{}".format(self.__all_log_path), 'a+') as f:
            f.write(message)
            f.write("\n")
        f.close()

    def SaveErrorLog(self, message):
        with open(r"{}".format(self.__error_log_path), 'a+') as f:
            f.write(message)
            f.write("\n")
        f.close()

    def SaveSendErrorLog(self, message):
        with open(r"{}".format(self.__send_error_log_path), 'a+') as f:
            f.write(message)
            f.write("\n")
        f.close()

    def SaveSendDoneLog(self, message):
        with open(r"{}".format(self.__send_done_log_path), 'a+') as f:
            f.write(message)
            f.write("\n")
        f.close()

    def tips(self, message):
        now_time_detail = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(now_time_detail + ("\033[34m [TIPS]: {}\033[0m").format(message))
        self.SaveAllLog(now_time_detail + (" [TIPS]: {}").format(message))

    def info(self, message):
        now_time_detail = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(now_time_detail + (" [INFO]: {}").format(message))
        self.SaveAllLog(now_time_detail + (" [INFO]: {}").format(message))

    def warning(self, message):
        now_time_detail = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(now_time_detail + ("\033[33m [WARNING]: {}\033[0m").format(message))
        self.SaveAllLog(now_time_detail + (" [WARNING]: {}").format(message))

    def error(self, message):
        now_time_detail = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(now_time_detail + ("\033[31m [ERROR]: {}\033[0m").format(message))
        self.SaveAllLog(now_time_detail + (" [ERROR]: {}").format(message))
        self.SaveErrorLog(now_time_detail + (" [ERROR]: {}").format(message))

    def send_error(self, message):
        self.SaveSendErrorLog(("{}").format(message))

    def done(self, message):
        now_time_detail = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(now_time_detail + ("\033[32m [DONE]: {}\033[0m").format(message))
        self.SaveAllLog(now_time_detail + (" [DONE]: {}").format(message))

    def send_done(self, message):
        self.SaveSendDoneLog(("{}").format(message))


Log = Log()
Lines = []


class EmailSender:
    def __init__(self, Subject, From, Content, smtp_user, smtp_passwd, smtp_server, Server=None, email_list=None,
                 file=None, sleep=0.5):
        self.Subject = Subject
        self.From = From
        self.Content = Content
        self.file = file
        self.smtp_user = smtp_user
        self.smtp_passwd = smtp_passwd
        self.smtp_server = smtp_server
        self.email_list = email_list
        self.Server = Server
        self.sleep = sleep

    def Sender(self):
        global Lines, client
        if not os.path.exists(r"{}".format(self.email_list)):
            Log.error('收件人列表文件未找到！')
            return
        File_Read(self.email_list)  # 读取全部内容 ，并以列表方式返回

        for line in Lines:
            rcptto = []
            rcptto.append(line.rstrip("\n"))
            self.victim = line.split('@')[0]
            # 显示的Cc收信地址
            rcptcc = []
            # Bcc收信地址，密送人不会显示在邮件上，但可以收到邮件
            rcptbcc = []
            # 全部收信地址，包含抄送地址，单次发送不能超过60人
            receivers = rcptto + rcptcc + rcptbcc

            # 参数判断
            if self.Subject == None or self.Subject == "":
                Log.error('邮件主题不存在！')
                return
            if self.smtp_user == None or self.smtp_user == "":
                Log.error('SMTP账号不存在！')
                return
            if self.smtp_passwd == None or self.smtp_passwd == "":
                Log.error('SMTP密码不存在！')
                return

                # 构建alternative结构
            msg = MIMEMultipart('alternative')
            msg['Subject'] = Header(self.Subject)
            if self.From[0] == None or self.From[0] == '':
                Log.error('发件人名称不存在！')
                return
            if self.From[1] == None or self.From[1] == '':
                self.From[1] = self.smtp_user
            msg['From'] = formataddr(self.From)  # 昵称+发信地址(或代发)
            # list转为字符串
            msg['To'] = ",".join(rcptto)
            msg['Cc'] = ",".join(rcptcc)
            # 自定义的回信地址，与控制台设置的无关。邮件推送发信地址不收信，收信人回信时会自动跳转到设置好的回信地址。
            # msg['Reply-to'] = replyto
            msg['Message-id'] = email.utils.make_msgid()
            msg['Date'] = email.utils.formatdate()

            email_content = self.Content.replace("victim", "victim=" + self.victim)

            # 加载远程图片以标记打开邮件的受害者
            if self.Server != ':':
                # Log.tips("监听地址: {}".format(self.Server))
                listen_code = '<img src="http://' + self.Server + '/click.php?victim=' + self.victim + '" style="display:none;" border="0">'
            else:
                listen_code = ''

            # 构建alternative的text/html部分
            texthtml = MIMEText(
                email_content + listen_code,
                _subtype='html', _charset='UTF-8')
            msg.attach(texthtml)

            # 附件
            if self.file != None and self.file != "":
                files = [r'{}'.format(self.file)]
                for t in files:
                    part_attach1 = MIMEApplication(open(t, 'rb').read())  # 打开附件
                    part_attach1.add_header('Content-Disposition', 'attachment', filename=t.rsplit('/', 1)[1])  # 为附件命名
                    msg.attach(part_attach1)  # 添加附件

            # 发送邮件
            try:
                client = smtplib.SMTP(self.smtp_server, 25)
                client.login(self.smtp_user, self.smtp_passwd)
                client.sendmail(self.smtp_user, receivers, msg.as_string())  # 支持多个收件人，最多60个
                client.quit()
                Log.done("{:<30}\t邮件发送成功！".format(rcptto[0]))
                Log.send_done("{}".format(rcptto[0]))
                time.sleep(self.sleep)
            except Exception as e:
                Log.error('邮件发送失败, {}'.format(e))
                Log.send_error(rcptto[0])


if __name__ == '__main__':
    mail_text_path = "./email/email.html" #邮件正文
    mail_to_list = "./target/email.txt" #收件人列表
    Log.tips("执行邮件正文读取操作！")
    f = open(mail_text_path, "r", encoding='utf-8')
    Log.done("邮件正文读取完成！")
    body_text = f.read()
    Log.tips("执行邮件发送操作！")
    E = EmailSender("【重要】数据迁移的通知", ['信息安全部', 'admin@admin.com'], body_text, "admin@admin.com", "pass",
                    "smtp.admin.com", "html.admin.com:8080", mail_to_list)
    E.Sender()
    Log.done("邮件发送完成!")

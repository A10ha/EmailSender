import base64
import smtplib
from tkinter.ttk import Combobox

import email
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from email.parser import Parser
import time
from tkinter import *
import tkinter.filedialog
import os
from datetime import datetime
import sys
import icon
import re

class StdoutRedirector(object):
    # 重定向输出类
    def __init__(self, text_widget):
        self.text_space = text_widget
        # 将其备份
        self.stdoutbak = sys.stdout
        self.stderrbak = sys.stderr

    def write(self, str):
        self.text_space.tag_config('fc_info', foreground='white')
        self.text_space.tag_config('fc_error', foreground='red')
        self.text_space.tag_config('fc_done', foreground='green')
        self.text_space.tag_config('fc_tips', foreground='cyan')
        if 'INFO' in str:
            self.text_space.insert(END, str, 'fc_info')
            self.text_space.insert(END, '\n')
        elif 'ERROR' in str:
            self.text_space.insert(END, str, 'fc_error')
            self.text_space.insert(END, '\n')
        elif 'DONE' in str:
            self.text_space.insert(END, str, 'fc_done')
            self.text_space.insert(END, '\n')
        elif 'TIPS' in str:
            self.text_space.insert(END, str, 'fc_tips')
            self.text_space.insert(END, '\n')
        self.text_space.see(END)
        self.text_space.update()

    def restoreStd(self):
        # 恢复标准输出
        sys.stdout = self.stdoutbak
        sys.stderr = self.stderrbak

    def flush(self):
        # 关闭程序时会调用flush刷新缓冲区，没有该函数关闭时会报错
        pass


# cur_path = os.path.dirname(os.path.realpath(__file__))  # 当前项目路径
cur_path = os.path.dirname(sys.executable)  # 打包运行路径
log_path = cur_path + '\\logs'  # log_path为存放日志的路径
if not os.path.exists(log_path): os.mkdir(log_path)  # 若不存在logs文件夹，则自动创建


class Log:

    def __init__(self):
        self.__now_time_detail = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # 当前日期格式化
        self.__now_time = datetime.now().strftime('%Y-%m-%d')
        self.__all_log_path = os.path.join(log_path, self.__now_time + "-all" + ".log")  # 收集所有日志信息文件
        self.__error_log_path = os.path.join(log_path, self.__now_time + "-error" + ".log")  # 收集错误日志信息文件

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

    def tips(self, message):
        print(self.__now_time + (" [TIPS]: {}").format(message))
        self.SaveAllLog(self.__now_time_detail + (" [TIPS]: {}").format(message))

    def info(self, message):
        print(self.__now_time + (" [INFO]: {}").format(message))
        self.SaveAllLog(self.__now_time_detail + (" [INFO]: {}").format(message))

    def warning(self, message):
        print(self.__now_time + (" [WARNING]: {}").format(message))
        self.SaveAllLog(self.__now_time_detail + (" [WARNING]: {}").format(message))

    def error(self, message):
        print(self.__now_time + (" [ERROR]: {}").format(message))
        self.SaveAllLog(self.__now_time_detail + (" [ERROR]: {}").format(message))
        self.SaveErrorLog(self.__now_time_detail + (" [ERROR]: {}").format(message))

    def done(self, message):
        print(self.__now_time + (" [DONE]: {}").format(message))
        self.SaveAllLog(self.__now_time_detail + (" [ERROR]: {}").format(message))


Log = Log()
stop_flag = 0
start_flag = 0
init_flag = 0
send_index = 0


def StopSend():
    global stop_flag, start_flag
    stop_flag = 1
    start_flag = 0


def StartSend():
    global stop_flag, start_flag
    stop_flag = 0
    start_flag = 1


class EmailSender:
    def __init__(self, Subject, From, Content, smtp_user, smtp_passwd, smtp_server, Server=None, email_list=None,
                 test_user=None,
                 file=None, img=None, sleep=0.5):
        self.Subject = Subject
        self.From = From
        self.Content = Content
        self.file = file
        self.smtp_user = smtp_user
        self.smtp_passwd = smtp_passwd
        self.smtp_server = smtp_server
        self.email_list = email_list
        self.Server = Server
        self.test_user = test_user
        self.img = img
        self.sleep = sleep

    def Sender(self):
        global client
        global stop_flag, start_flag
        global send_index
        if self.test_user != None and self.test_user != '':
            lines = [self.test_user]
        else:
            if not os.path.exists(r"{}".format(self.email_list)):
                Log.error('收件人列表文件未找到！')
                return False
            f = open(r"{}".format(self.email_list), "r")
            lines = f.readlines()  # 读取全部内容 ，并以列表方式返回

        for line in lines[send_index:]:
            if stop_flag == 1:
                Log.tips("暂停发送，目前成功发送: {} 封".format(send_index))
                Log.tips("发送截至目标: {}".format(line))
                break
            if line == "\n" or line == '':
                continue
            rcptto = []
            rcptto.append(line.rstrip("\n"))
            self.victim = line.split('@')[0]
            # 显示的Cc收信地址
            rcptcc = []
            # Bcc收信地址，密送人不会显示在邮件上，但可以收到邮件
            rcptbcc = []
            # 全部收信地址，包含抄送地址，单次发送不能超过60人
            receivers = rcptto + rcptcc + rcptbcc

            # 构建alternative结构
            msg = MIMEMultipart('alternative')
            msg['Subject'] = Header(self.Subject)
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

            email_content = self.Content

            # base64加载图片
            if self.img != None and self.img != '':
                with open(r"{}".format(self.img), "rb") as i:
                    b64str = base64.b64encode(i.read())
                # Log.info(str(b64str))
                img_code = '<img src="data:image/png;base64,' + b64str.decode('utf-8') + '" />'
            else:
                img_code = ''

            # 加载远程图片以标记打开邮件的受害者
            if self.Server != ':':
                Log.tips("监听地址: {}".format(self.Server))
                listen_code = '<img src="http://' + self.Server + '/' + self.victim + '" style="display:none;" border="0">'
            else:
                listen_code = ''

            # 构建alternative的text/html部分
            texthtml = MIMEText(
                email_content + img_code + listen_code,
                _subtype='html', _charset='UTF-8')
            msg.attach(texthtml)

            # 附件
            if self.file != None and self.file != "":
                files = [r'{}'.format(self.file)]
                for t in files:
                    part_attach1 = MIMEApplication(open(t, 'rb').read())  # 打开附件
                    # print(t.rsplit('/', 1)[1])
                    part_attach1.add_header('Content-Disposition', 'attachment', filename=t.rsplit('/', 1)[1])  # 为附件命名
                    msg.attach(part_attach1)  # 添加附件

            # 发送邮件
            try:
                # 判断长度，当测试单次发送时值为“腾讯企业邮:*25”长度为9，当批量发送时，经过strip('*:25')，长度为5
                if len(self.smtp_server) == 9 or len(self.smtp_server) == 5:
                    self.smtp_server = self.smtp_server.strip('*:25')
                    if self.smtp_server == "腾讯企业邮":
                        # 若需要加密使用SSL，可以这样创建client
                        client = smtplib.SMTP_SSL('smtp.exmail.qq.com', 465)
                    elif self.smtp_server == "网易企业邮":
                        client = smtplib.SMTP_SSL('smtphz.qiye.163.com', 465)
                    elif self.smtp_server == "阿里企业邮":
                        client = smtplib.SMTP_SSL('smtpdm.aliyun.com', 465)
                else:
                    tmp = self.smtp_server.split('*')[1]
                    server_ip = tmp.split(':')[0]
                    server_port = tmp.split(':')[1]
                    if int(server_port) == 465:
                        client = smtplib.SMTP_SSL(server_ip, int(server_port))
                    else:
                        client = smtplib.SMTP(server_ip, int(server_port))
                # 开启DEBUG模式
                # client.set_debuglevel(0)
                # 发件人和认证地址必须一致
                client.login(self.smtp_user, self.smtp_passwd)
                # 备注：若想取到DATA命令返回值,可参考smtplib的sendmail封装方法:
                # 使用SMTP.mail/SMTP.rcpt/SMTP.data方法
                # Log.tips("执行发送")
                client.sendmail(self.smtp_user, receivers, msg.as_string())  # 支持多个收件人，最多60个
                client.quit()
                Log.done('{:<30}\t邮件发送成功！'.format(rcptto[0]))
                send_index += 1
                time.sleep(self.sleep)
            except smtplib.SMTPConnectError as e:
                Log.error('邮件发送失败，连接失败: [Code]{}，[Error]{}'.format(e.smtp_code, e.smtp_error))
            except smtplib.SMTPAuthenticationError as e:
                Log.error('邮件发送失败，认证错误: [Code]{}，[Error]{}'.format(e.smtp_code, e.smtp_error))
            except smtplib.SMTPSenderRefused as e:
                Log.error('邮件发送失败，发件人被拒绝: [Code]{}，[Error]{}'.format(e.smtp_code, e.smtp_error))
            except smtplib.SMTPRecipientsRefused as e:
                Log.error('邮件发送失败，收件人被拒绝: [Error]{}'.format(e))
            except smtplib.SMTPDataError as e:
                Log.error('邮件发送失败，数据接收拒绝:[Code]{}，[Error]{}'.format(e.smtp_code, e.smtp_error))
            except smtplib.SMTPException as e:
                Log.error('邮件发送失败, {}'.format(str(e)))
            except Exception as e:
                Log.error('邮件发送异常, {}'.format(str(e)))


def read_mail(path):
    if os.path.exists(path):
        with open(path) as fp:
            email = fp.read()
            return email
    else:
        Log.error('EML模板文件不存在！')
        return None


def emailInfo(emailpath):
    raw_email = read_mail(emailpath)  # 将邮件读到一个字符串里面
    if raw_email == None:
        return None
    # print('EmailPath : ', emailpath)
    emailcontent = Parser().parsestr(raw_email)  # 经过parsestr处理过后生成一个字典
    # for k, v in emailcontent.items():
    #     print(k, v)
    From = emailcontent['From']
    To = emailcontent['To']
    Subject = emailcontent['Subject']
    Date = emailcontent['Date']
    MessageID = emailcontent['Message-ID']
    XOriginatingIP = emailcontent['X-Originating-IP']

    User = From.split("<")[0].strip()
    if "<" in From:
        From = re.findall(".*<(.*)>.*", From)[0]
    if "<" in To:
        To = re.findall(".*<(.*)>.*", To)[0]
    if "<" in MessageID:
        MessageID = re.findall(".*<(.*)>.*", MessageID)[0]
    try:
        Subject = base64.b64decode(Subject.split("?")[3]).decode('utf-8')
    except Exception:
        pass
    try:
        User = base64.b64decode(User.split("?")[3]).decode('utf-8')
    except Exception:
        pass

    Log.info("[From-User]:\t{}".format(User))
    Log.info("[From]:\t{}".format(From))
    Log.info("[X-Originating-IP]:\t{}".format(XOriginatingIP))
    Log.info("[To]:\t{}".format(To))
    Log.info("[Subject]:\t{}".format(Subject))
    Log.info("[Message-ID]:\t{}".format(MessageID))
    Log.info("[Date]:\t{}".format(Date))

    # 循环信件中的每一个mime的数据块
    for par in emailcontent.walk():
        if not par.is_multipart():  # 这里要判断是否是multipart（无用数据）
            content = par.get_payload(decode=True)
            try:
                if content.decode('utf-8', 'ignore').startswith('<'):
                    # print("content:\t\n", content.decode('utf-8', 'ignore').strip())  # 解码出文本内容，直接输出来就可以了。
                    return content.decode('utf-8', 'ignore').strip()
            except UnicodeDecodeError:
                if content.decode('gbk', 'ignore').startswith('<'):
                    # print("content:\t\n", content.decode('utf-8', 'ignore').strip())  # 解码出文本内容，直接输出来就可以了。
                    return content.decode('gbk', 'ignore').strip()


class EmailSenderGUI():
    def __init__(self, init_window):
        self.init_window = init_window

    def set_init_windows(self):
        self.init_window.title("钓鱼邮件发送工具")
        # self.init_window.geometry('1230x810') # 2k窗口大小，若1k分辨率，请自己调试
        self.init_window.geometry('1400x1000')  # 4k窗口大小
        with open('tmp.ico', 'wb') as tmp:
            tmp.write(base64.b64decode(icon.Icon().img))
        root.iconbitmap('tmp.ico')
        os.remove('tmp.ico')
        self.init_window.resizable(0, 0)
        # pop = Pmw.Balloon(self.init_window)
        # 基础配置项
        self.basic_setting_label = Label(self.init_window, text="基础配置"
                                         )
        self.basic_setting_label.grid(row=0, column=0)
        self.basic_setting = Frame(self.init_window)
        # SMTP服务器选择
        self.smtp_server_label = Label(self.basic_setting, text='SMTP服务器:')
        self.smtp_server_label.grid(row=0, column=0)
        self.smtp_server_combobox = Combobox(self.basic_setting, width=25)
        self.smtp_server_combobox.grid(row=0, column=1, sticky='w', pady=3)
        self.smtp_server_combobox['value'] = ('腾讯企业邮', '网易企业邮', '阿里企业邮')
        self.smtp_server_combobox.current(0)
        self.send_sleep_time_label = Label(self.basic_setting, text="发送间隔(s):")
        self.send_sleep_time_label.grid(row=0, column=1, sticky='e', padx=48)
        self.send_sleep_time_entry = Entry(self.basic_setting, width=6)
        self.send_sleep_time_entry.insert(tkinter.INSERT, '0.5')
        self.send_sleep_time_entry.grid(row=0, column=1, sticky='e', pady=3)
        # 账号输入框
        self.smtp_account_label = Label(self.basic_setting, text="SMTP账号:")
        self.smtp_account_label.grid(row=1, column=0)
        self.smtp_account_entry = Entry(self.basic_setting, width=45)
        # pop.bind(self.smtp_account_label, "必填")
        self.smtp_account_entry.grid(row=1, column=1, pady=3)
        # 密码输入框
        self.smtp_passwd_label = Label(self.basic_setting, text="SMTP密码:")
        self.smtp_passwd_label.grid(row=2, column=0)
        self.smtp_passwd_entry = Entry(self.basic_setting, show='*', width=45)
        # pop.bind(self.smtp_passwd_label, "必填")
        self.smtp_passwd_entry.grid(row=2, column=1, pady=3)
        # 自定义邮件服务器
        self.smtp_server_self_label = Label(self.basic_setting, text="自建SMTP:")
        self.smtp_server_self_label.grid(row=3, column=0)
        self.smtp_server_self_entry = Entry(self.basic_setting, width=30)
        # pop.bind(self.smtp_server_self_entry, "当设置自定邮服时，SMTP选择无效，填写格式：127.0.0.1:465")
        self.smtp_server_self_entry.grid(row=3, column=1, sticky='w')
        self.smtp_server_self_port_entry = Entry(self.basic_setting, width=8)
        self.smtp_server_self_port_entry.insert(tkinter.INSERT, '25')
        self.smtp_server_self_blank = Label(self.basic_setting, text="PORT: ", width=5)
        self.smtp_server_self_blank.grid(row=3, column=1, sticky='e', padx=60, pady=3)
        self.smtp_server_self_port_entry.grid(row=3, column=1, sticky='e', pady=3)
        # 监听服务器地址
        self.listen_server_label = Label(self.basic_setting, text="监听服务器地址:")
        self.listen_server_label.grid(row=4, column=0)
        self.listen_server_entry = Entry(self.basic_setting, width=30)
        # pop.bind(self.listen_server_entry, "填写格式：127.0.0.1:8080")
        self.listen_server_entry.grid(row=4, column=1, sticky='w', pady=3)
        self.listen_server_blank = Label(self.basic_setting, text="PORT: ", width=5)
        self.listen_server_blank.grid(row=4, column=1, sticky='e', padx=60)
        self.listen_server_port_entry = Entry(self.basic_setting, width=8)
        self.listen_server_port_entry.grid(row=4, column=1, sticky='e', pady=3)

        # 测试发送
        self.test_user_label = Label(self.basic_setting, text="测试接收邮箱:")
        self.test_user_label.grid(row=5, column=0)
        self.test_user_entry = Entry(self.basic_setting, width=45)
        self.test_user_entry.grid(row=5, column=1, pady=3, )
        self.test_email_sender = Button(self.basic_setting, text="测试发送", command=lambda: self.SenderHandler(0),
                                        height=1)
        self.test_email_sender.grid(row=5, column=2, padx=3)
        self.basic_setting.grid(row=1, column=0)

        # 邮件配置
        self.email_setting_label = Label(self.init_window, text="邮件配置")
        self.email_setting_label.grid(row=2, column=0)
        self.email_setting = Frame(self.init_window)
        # 邮件头配置
        # 伪造发件人配置
        self.email_sender_name_label = Label(self.email_setting, text="伪造发件人名称:")
        self.email_sender_name_label.grid(row=0, column=0)
        self.email_sender_name_entry = Entry(self.email_setting, width=45)
        # pop.bind(self.email_sender_name_entry, "必填")
        self.email_sender_name_entry.grid(row=0, column=1, pady=3)
        self.email_sender_email_label = Label(self.email_setting, text="伪造发件人邮箱:")
        self.email_sender_email_label.grid(row=1, column=0)
        self.email_sender_email_entry = Entry(self.email_setting, width=45)
        self.email_sender_email_entry.grid(row=1, column=1, pady=3)
        # 邮件主题
        self.email_subject_label = Label(self.email_setting, text="邮件主题:")
        self.email_subject_label.grid(row=2, column=0)
        self.email_subject_entry = Entry(self.email_setting, width=45)
        # pop.bind(self.email_subject_entry, "必填")
        self.email_subject_entry.grid(row=2, column=1, pady=3)

        # 邮件原文EML文件选择
        self.eml_file_label = Label(self.email_setting, text="EML文件:")
        self.eml_file_label.grid(row=3, column=0)
        self.EML_file = StringVar()
        self.eml_file_entry = Entry(self.email_setting, textvariable=self.EML_file, width=45)
        self.eml_file_entry.grid(row=3, column=1, pady=3)
        self.eml_file_selector = Button(self.email_setting, text="选择文件",
                                        command=lambda: self.FileSelector(self.eml_file_entry, self.EML_file, 1))
        self.eml_file_selector.grid(row=3, column=2, padx=5)

        # 收件人列表
        self.email_label = Label(self.email_setting, text="收件人列表:")
        self.email_label.grid(row=4, column=0)
        self.EMAIL_list = StringVar()
        self.email_list_entry = Entry(self.email_setting, textvariable=self.EMAIL_list, width=45)
        self.email_list_entry.grid(row=4, column=1, pady=3)
        self.email_file_selector = Button(self.email_setting, text="选择文件",
                                          command=lambda: self.FileSelector(self.email_list_entry, self.EMAIL_list, 0))
        self.email_file_selector.grid(row=4, column=2, padx=5)

        # 邮件附件选择
        self.file_label = Label(self.email_setting, text="邮件附件:")
        self.file_label.grid(row=5, column=0)
        self.EMAIL_file = StringVar()
        self.file_entry = Entry(self.email_setting, textvariable=self.EMAIL_file, width=45)
        self.file_entry.grid(row=5, column=1, pady=3)
        self.file_selector = Button(self.email_setting, text="选择附件",
                                    command=lambda: self.FileSelector(self.file_entry, self.EMAIL_file, 0))
        self.file_selector.grid(row=5, column=2, padx=5)

        # 邮件正文图片插入
        self.img_label = Label(self.email_setting, text="正文图片:")
        self.img_label.grid(row=6, column=0)
        self.EMAIL_img = StringVar()
        self.img_entry = Entry(self.email_setting, textvariable=self.EMAIL_img, width=45)
        self.img_entry.grid(row=6, column=1, pady=3)
        self.img_selector = Button(self.email_setting, text="选择图片",
                                   command=lambda: self.FileSelector(self.img_entry, self.EMAIL_img, 0))
        self.img_selector.grid(row=6, column=2, padx=5)
        self.email_setting.grid(row=3, column=0)

        # 邮件编辑区头部
        self.email_editor_label = Label(self.init_window, text="邮件正文编辑")
        self.email_editor_label.grid(row=0, column=1)
        # 邮件编辑区
        self.email_editor = Frame(self.init_window)
        self.email_editor_text = Text(self.email_editor, height=56, width=90)
        self.email_editor_text.grid(row=0, column=0)
        self.email_editor.grid(row=1, column=1, rowspan=5)
        self.email_editor_text.tag_configure("found", background="yellow")

        # 正文搜索
        self.email_function_frame = Frame(self.init_window)
        self.email_searcher_entry = Entry(self.email_function_frame, width=35)
        self.email_searcher_entry.grid(row=0, column=0, pady=3)
        self.email_searcher = Button(self.email_function_frame, text="搜索", command=self.TextSearcher)
        self.email_searcher.grid(row=0, column=1, padx=5)

        # 邮件按钮
        self.email_preview = Button(self.email_function_frame, text="邮件预览", command=self.HTMLRunner)
        self.email_preview.grid(row=0, column=2, padx=5)
        # 邮件发送
        self.email_sender = Button(self.email_function_frame, text="批量发送", command=lambda: self.SenderHandler(1))
        self.email_sender.grid(row=0, column=3, padx=5)
        # 重置
        self.email_reset_sender = Button(self.email_function_frame, text="重置", command=lambda: self.Reset())
        self.email_reset_sender.grid(row=0, column=4, padx=5)
        self.email_function_frame.grid(column=1)

        # 控制台输出
        self.log_print_label = Label(self.init_window, text="控制台输出")
        self.log_print_label.grid(row=4, column=0)
        self.log_print_from = Frame(self.init_window)
        self.log_print_window = Text(self.log_print_from, bg='black')
        sys.stdout = StdoutRedirector(self.log_print_window)
        self.log_print_window.grid(row=0, column=0)
        self.log_print_from.grid(row=5, column=0, padx=10)

        # banner区
        self.banner_frame = Frame(self.init_window)
        self.banner_label = Label(self.banner_frame, text="Github: https://github.com/A10ha ")
        self.banner_label.grid(row=0, column=0)
        self.banner_frame.grid(row=6, column=0, pady=20)

    def TextSearcher(self):
        self.email_editor_text.tag_remove("found", "1.0", END)
        start = "1.0"
        key = self.email_searcher_entry.get()

        if (len(key.strip()) == 0):
            return
        while True:
            pos = self.email_editor_text.search(key, start, END)
            # print("pos= ",pos) # pos=  3.0  pos=  4.0  pos=
            if (pos == ""):
                break
            self.email_editor_text.tag_add("found", pos, "%s+%dc" % (pos, len(key)))
            start = "%s+%dc" % (pos, len(key))

    def FileSelector(self, File_Entry, Entry_Value, flag):
        Entry_Value.set('')
        self.filename = tkinter.filedialog.askopenfilename()
        if self.filename != '':
            File_Entry.insert(tkinter.INSERT, self.filename)
            if flag == 1:
                self.HTMLReader(self.filename)

    def HTMLRunner(self):
        HTML_content = self.email_editor_text.get("0.0", "end")
        if self.img_entry.get() != None and self.img_entry.get() != '':
            with open(r"{}".format(self.img_entry.get()), "rb") as i:
                b64str = base64.b64encode(i.read())
            img_code = '<img src="data:image/png;base64,' + b64str.decode(
                'utf-8') + '" style="display:block;margin:0 auto;" />'
        else:
            img_code = ''
        # self.email_editor_text.insert(tkinter.INSERT, img_code)
        with open('temp.html', 'wb') as f:
            HTML = HTML_content + img_code
            f.write(HTML.encode('utf-8'))
        os.popen('start temp.html')

    def HTMLReader(self, eml_file):
        self.email_editor_text.delete('1.0', 'end')
        Log.tips("============================================================")
        HTML_content = emailInfo(eml_file)
        self.email_editor_text.insert(tkinter.INSERT, HTML_content)
        with open('temp.html', 'wb') as f:
            f.write(HTML_content.encode('utf-8'))

    def SenderHandler(self, flag):
        global start_flag, stop_flag, init_flag, E
        if init_flag == 0:
            self.InitSender(flag)
            return
        if start_flag == 1:
            StopSend()
            self.email_sender['text'] = "恢复发送"
            self.email_reset_sender['state'] = 'normal'
        else:
            StartSend()
            self.email_sender['text'] = "暂停发送"
            self.email_reset_sender['state'] = 'disabled'
            E.Sender()

    def Reset(self):
        global init_flag, start_flag, stop_flag, send_index
        init_flag = 0
        start_flag = 0
        stop_flag = 0
        send_index = 0
        time.sleep(0.5)
        Log.tips("发送队列已重置，可重新选择收件人邮箱文件进行重新发送。\n")
        self.email_sender['text'] = "批量发送"
        self.test_email_sender['state'] = 'normal'
        self.email_editor_text['state'] = 'normal'
        self.send_sleep_time_entry['state'] = 'normal'
        self.email_sender_name_entry['state'] = 'normal'
        self.email_sender_email_entry['state'] = 'normal'
        self.email_subject_entry['state'] = 'normal'

    def InitSender(self, flag):
        global init_flag, start_flag, send_index, E
        if self.smtp_server_self_entry.get() == None or self.smtp_server_self_entry.get() == '':
            Log.tips("SMTP服务器: {}".format(self.smtp_server_combobox.get()))
        else:
            Log.tips(
                "SMTP服务器: {}".format(self.smtp_server_self_entry.get() + ':' + self.smtp_server_self_port_entry.get()))
        # 参数判断
        if self.email_subject_entry.get() == None or self.email_subject_entry.get() == "":
            Log.error('邮件主题不存在！')
            return False
        if self.smtp_account_entry.get() == None or self.smtp_account_entry.get() == "":
            Log.error('SMTP账号不存在！')
            return False
        if self.smtp_passwd_entry.get() == None or self.smtp_passwd_entry.get() == "":
            Log.error('SMTP密码不存在！')
            return False
        if self.email_sender_name_entry.get() == None or self.email_sender_name_entry.get() == '':
            Log.error('发件人名称不存在！')
            return False
        if flag == 0:
            if self.test_user_entry.get() == None or self.test_user_entry.get() == '':
                Log.error("测试收件人邮箱未设置！")
            else:
                Log.tips("测试发送: {}".format(self.test_user_entry.get()))
                E = EmailSender(self.email_subject_entry.get(),
                                [self.email_sender_name_entry.get(), self.email_sender_email_entry.get()],
                                self.email_editor_text.get("0.0", "end"),
                                self.smtp_account_entry.get(), self.smtp_passwd_entry.get(),
                                self.smtp_server_combobox.get() + '*' + self.smtp_server_self_entry.get() + ':' + self.smtp_server_self_port_entry.get(),
                                self.listen_server_entry.get() + ':' + self.listen_server_port_entry.get(),
                                None,
                                self.test_user_entry.get(),
                                self.file_entry.get(), self.img_entry.get(), float(self.send_sleep_time_entry.get()))
                send_index = 0
                E.Sender()
        elif flag == 1:
            if self.email_list_entry.get() == None or self.email_list_entry.get() == '':
                Log.error("收件人列表文件路径未设置！")
            else:
                init_flag = 1
                start_flag = 1
                self.email_sender['text'] = "暂停发送"
                self.email_reset_sender['state'] = 'disabled'
                self.test_email_sender['state'] = 'disabled'
                self.email_editor_text['state'] = 'disabled'
                self.send_sleep_time_entry['state'] = 'disabled'
                self.email_sender_name_entry['state'] = 'disabled'
                self.email_sender_email_entry['state'] = 'disabled'
                self.email_subject_entry['state'] = 'disabled'
                Log.tips("批量发送: {}".format(self.email_list_entry.get()))
                E = EmailSender(self.email_subject_entry.get(),
                                [self.email_sender_name_entry.get(), self.email_sender_email_entry.get()],
                                self.email_editor_text.get("0.0", "end"),
                                self.smtp_account_entry.get(), self.smtp_passwd_entry.get(),
                                self.smtp_server_combobox.get() + '*' + self.smtp_server_self_entry.get() + ':' + self.smtp_server_self_port_entry.get(),
                                self.listen_server_entry.get() + ':' + self.listen_server_port_entry.get(),
                                self.email_list_entry.get(),
                                None,
                                self.file_entry.get(), self.img_entry.get(), float(self.send_sleep_time_entry.get()))
                E.Sender()


if __name__ == '__main__':
    root = Tk()  # 实例化出一个父窗口
    PORTAL = EmailSenderGUI(root)
    # 设置根窗口默认属性
    PORTAL.set_init_windows()
    # 父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示
    root.mainloop()

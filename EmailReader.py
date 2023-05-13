import poplib
from email.parser import Parser
import re


#退件邮件收件人获取
#连接邮箱服务器
server = poplib.POP3('smtp.xxxx.com')
server.user('xxxx')
server.pass_('xxxx')

#获取邮件列表
resp, mails, octets = server.list()

#指定要读取的邮件主题
subject = 'Undelivered Mail Returned to Sender'

#获取邮件的字符编码，首先在message中寻找编码，如果没有，就在header的Content-Type中寻找
def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos+8:].strip()
    return charset

def get_content(msg):
    for part in msg.walk():
        content_type = part.get_content_type()
        charset = guess_charset(part)
        #如果有附件，则直接跳过
        if part.get_filename()!=None:
            continue
        email_content_type = ''
        content = ''
        if content_type == 'text/plain':
            email_content_type = 'text'
        elif content_type == 'text/html':
            continue #不要html格式的邮件
        if charset:
            try:
                content = part.get_payload(decode=True).decode(charset)
            #这里遇到了几种由广告等不满足需求的邮件遇到的错误，直接跳过了
            except AttributeError:
                print('type error')
            except LookupError:
                print("unknown encoding: utf-8")
        if email_content_type =='':
            continue
            #如果内容为空，也跳过
        return content
        #邮件的正文内容就在content中


if __name__ == '__main__':
    f = open("reject.txt", 'a+')
    #遍历邮件列表
    for i in range(len(mails), 0, -1):
        resp, lines, octets = server.retr(i)
        #合并邮件内容
        message = b'\n'.join(lines).decode('utf-8')
        # 将邮件内容解析成Message对象
        emailcontent = Parser().parsestr(message)
        #如果主题匹配，则输出邮件内容
        if subject == emailcontent['Subject'] :
            content = get_content(emailcontent)
            match = re.findall(r'<.*?>', content)
            result = match[0].lstrip('<').rstrip('>')
            print(result)
            f.write(result)
            f.write("\n")
            # break
    f.close()


#关闭连接
server.quit()
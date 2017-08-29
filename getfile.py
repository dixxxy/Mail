#!/usr/bin/env python3
# -*- coding: GB18030 -*-

import poplib
import email
import os
import configparser


from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr


def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def guess_charset(msg):
    # 先从msg对象获取编码:
    charset = msg.get_charset()
    if charset is None:
        # 如果获取不到，再从Content-Type字段获取:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

def get_email_headers(msg):
    # 邮件的From, To, Subject存在于根对象上:
    headers = {}
    for header in ['From', 'To', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['date'] = value
            if header == 'Subject':
                # 需要解码Subject字符串:
                subject = decode_str(value)
                headers['subject'] = subject
            else:
                # 需要解码Email地址:
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                value = u'%s <%s>' % (name, addr)
                if header == 'From':
                    from_address = value
                    headers['from'] = from_address
                else:
                    to_address = value
                    headers['to'] = to_address
    content_type = msg.get_content_type()
    print('head content_type:%s '% (content_type))
    return headers

def category(msg):
    new_headers = get_email_headers(msg)
    new_from = []
    new_from = new_headers.get('from')
    return new_from
# indent用于缩进显示:
def get_email_content(message, base_save_path):
    j = 0
    content = ''
    attachment_files = []
    for part in message.walk():
        #message_headers = get_email_headers(message)
        j = j + 1
        file_name = part.get_filename()
        print(file_name)
        charset = part.get_charset()
        contentType = part.get_content_type()
        print('-----保存附件-----')
        if file_name:  # Attachment
            # Decode filename
            h = email.header.Header(file_name)
            print(h)
            dh = email.header.decode_header(h)
            print(dh)
            filename = dh[0][0]
            if dh[0][1]:  # 如果包含编码的格式，则按照该格式解码
                filename = str(filename, dh[0][1])
            data = part.get_payload(decode=True)

            print('文件名：%s' % (filename))
            print("当前工作目录 : %s" % os.getcwd())
            os.chdir(base_save_path)
            print("当前工作目录 : %s" % os.getcwd())
            print('文件名类型：%s' % type(filename))
            try:
                att_file = open(filename,'wb')
                attachment_files.append(filename)
            except:
                print('附件名中有非法字符，自动换一个')
                filename = decode_str(filename)
                print('文件名：%s' % (filename))
                print('文件名类型：%s' % type(filename))
                att_file = open(filename,'wb')
            try:
                att_file.write(data)
            except:
                print('文件内容有问题')
                data = b'part.get_payload(decode=True)'
                att_file.write(data)

            att_file.close()


        elif contentType == 'text/plain' or contentType == 'text/html':
            # 保存正文
            data = part.get_payload(decode=True)
            charset = guess_charset(part)
            if charset:
                charset = charset.strip().split(';')[0]
                print('charset:%s' % (charset))
                data = data.decode(charset, 'ignore')
            content = data
    return content, attachment_files


if __name__ == '__main__':
    config = configparser.ConfigParser()
    config.read('Users.ini')

    config.write(open('Users.ini','w'))

    emailaddress = config.get('CNNIC2','user') #填写邮箱地址
    print(emailaddress)
    password = config.get('CNNIC2','password') #填写邮箱密码
    print(password)
    pop3_server = config.get('CNNIC2','pop')   #邮箱服务器
    print(pop3_server)

    # 连接到POP3服务器:
    #server = poplib.POP3_SSL(pop3_server,port=995)      #如果要访问QQ邮箱时，需要这句
    server = poplib.POP3(pop3_server)
    # 可以打开或关闭调试信息:
    # server.set_debuglevel(1)
    # POP3服务器的欢迎文字:
    print(server.getwelcome())
    # 身份认证:
    server.user(emailaddress)
    server.pass_(password)
    # stat()返回邮件数量和占用空间:
    messagesCount, messagesSize = server.stat()
    print('messagesCount:%s' %  (messagesCount))
    print('messagesSize:%s'% (messagesSize))
    # list()返回所有邮件的编号:
    resp, mails, octets = server.list()
    print('------ resp ------')
    print(resp)  # +OK 46 964346 响应的状态 邮件数量 邮件占用的空间大小
    print('------ mails ------')
    print(mails)  # 所有邮件的编号及大小的编号list，['1 2211', '2 29908', ...]
    print('------ octets ------')
    print(octets)

    new_headers_from = []
    # 获取最新一封邮件, 注意索引号从1开始:
    length = len(mails)
    for i in range(length):
        resp, lines, octets = server.retr(i + 1)
        # lines存储了邮件的原始文本的每一行,
        # 可以获得整个邮件的原始文本:
        msg_content = b"\r\n".join(lines).decode('latin-1')
        # 把邮件内容解析为Message对象：
        msg = Parser().parsestr(msg_content)

        # 但是这个Message对象本身可能是一个MIMEMultipart对象，即包含嵌套的其他MIMEBase对象，
        # 嵌套可能还不止一层。所以我们要递归地打印出Message对象的层次结构：
        print('---------- 解析之后 ----------')
        msg_headers = get_email_headers(msg)
        new_from = category(msg)
        new_headers_from.append(new_from)

        print('subject:%s'% (msg_headers.get('subject')))
        print('from_address:%s'% (msg_headers.get('from')))
        print('to_address:%s'% (msg_headers.get('to')))
        print('date:%s'% (msg_headers.get('date')))

        title = new_headers_from[i]
        print(title)
        #cur_dir = r'G:/GetFile/cnnic/'         #新建存附件的文件夹 使用绝对路径
        cur_dir = config.get('CNNIC2','dir')
        new_title = title.split('<')
        n_title = new_title[0]

        if n_title.strip() == '':
            p_title = new_title[1].split('@')
            final_title = p_title[0]
        else:
            final_title = n_title
        print(final_title)
        new_path = os.path.join(cur_dir, final_title)
        if not os.path.isdir(new_path):
            os.makedirs(new_path)
        base_save_path = new_path
        print(base_save_path)
        content, attachment_files = get_email_content(msg, base_save_path)
        print('content:%s'% (content))
        print('attachment_files:%s ' %  (attachment_files))



    # 关闭连接:
    server.quit()
